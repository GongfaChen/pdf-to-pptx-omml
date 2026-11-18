#!/usr/bin/env python3
"""license_check.py — 许可证钩子（免费全通过版）。

设计目标（open-core 分层，见 README「License / 授权」一节）：
- 现阶段：无 license 文件 → 一切功能免费全通过（本仓库所有脚本静默放行）。
- 未来付费层：只需 ①生成密钥对并替换下方 PUBLIC_KEY_HEX（或设环境变量），
  ②新付费脚本调用 require_license("pipeline") 等付费 feature。
  老用户环境零破坏，切换当天即生效。

验证语义（退出码与 pdf-image-text 同风格）：
  0 通过（免费模式或有效 license；--verify 时用 status: FREE / OK 区分）
  1 未找到 license 文件（免费模式，非错误）／ 需要 license 的付费功能被拒
  2 签名无效
  3 已过期
  4 产品 / 功能范围不符
  5 license 格式损坏
  6 机器绑定不符（预留，未启用）
  7 内部错误（公钥缺失或格式错误）

license.json 格式：
  {
    "product": "pdf-to-pptx-omml",
    "features": ["probe", "core"],
    "issued_to": "user@example.com",
    "issued_at": "2026-11-18",
    "expires": "2027-11-18",
    "signature": "<ed25519 hex，对以上字段按 sort_keys+紧凑分隔符序列化后签名>"
  }

CLI 用法：
  python license_check.py --generate-keys [目录]   生成密钥对（私钥不要提交！）
  python license_check.py --issue --private-key 种子.hex --email a@b.c --features core,pipeline --days 365
  python license_check.py --verify [license.json]  校验一份 license（退出码见上表）
  python license_check.py --self-test              生成→签发→校验 全链路自检

依赖：仅标准库（hashlib / json / argparse / datetime）。Ed25519 为 RFC 8032
纯 Python 实现（点运算走仿射坐标，实测校验 152–163 ms/次，约 0.16 s；不做结果缓存——缓存文件
用户可写，等于给付费闸开后门，宁可慢 150 ms）。

安全模型（重要）：
- 发布构建只信任模块内 PUBLIC_KEY_HEX；环境变量公钥仅在开发模式
  （PDF2PPTX_OMML_DEV=1）生效，否则忽略并告警——否则用户可自签 license 绕过付费闸。
- 无 license 文件时：FREE_FEATURES 永久免费；PAID_FEATURES 直接拒绝（退出码 1）。
  现在 PAID_FEATURES 为空，等价于"全部免费"，上线付费层时只改这一个集合。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

# ---------------------------------------------------------------------------
# 部署配置：上线付费层时把占位公钥替换成 --generate-keys 产出的真实公钥
# ---------------------------------------------------------------------------
PUBLIC_KEY_HEX = "00" * 32  # ← 占位：全零公钥 = 未配置许可体系，一律免费放行
PUBLIC_KEY_ENV = "PDF2PPTX_OMML_PUBKEY"  # 仅在开发模式下生效（见 _dev_mode）
DEV_MODE_ENV = "PDF2PPTX_OMML_DEV"  # =1 时允许环境变量公钥，并禁用发布语义
LICENSE_ENV = "PDF2PPTX_OMML_LICENSE"  # 显式指定 license 文件路径
PRODUCT = "pdf-to-pptx-omml"

FREE_FEATURES = {"probe"}  # 永远免费，无论未来如何配置
# 上线付费层时把付费 feature 加进来（如 "pipeline"）；现在为空 = 全部免费。
# 无 license 文件时，PAID_FEATURES 里的 feature 会以退出码 1 拒绝，其余照旧免费。
PAID_FEATURES: frozenset[str] = frozenset()
PLACEHOLDER_PUBKEY = "00" * 32

# ---------------------------------------------------------------------------
# Ed25519（RFC 8032，纯 Python 实现）
# ---------------------------------------------------------------------------
_Q = 2**255 - 19
_L = 2**252 + 27742317777372353535851937790883648493
_D = (-121665 * pow(121666, _Q - 2, _Q)) % _Q
_I = pow(2, (_Q - 1) // 4, _Q)


def _inv(x: int) -> int:
    return pow(x, _Q - 2, _Q)


def _xrecover(y: int) -> int:
    xx = (y * y - 1) * _inv(_D * y * y + 1)
    x = pow(xx, (_Q + 3) // 8, _Q)
    if (x * x - xx) % _Q != 0:
        x = (x * _I) % _Q
    if x % 2 != 0:
        x = _Q - x
    return x


_By = (4 * _inv(5)) % _Q
_B = (_xrecover(_By) % _Q, _By % _Q)
_IDENT = (0, 1)


def _on_curve(p: tuple[int, int]) -> bool:
    x, y = p
    return (-x * x + y * y - 1 - _D * x * x * y * y) % _Q == 0


def _edwards(p1, p2):
    x1, y1 = p1
    x2, y2 = p2
    x3 = (x1 * y2 + x2 * y1) * _inv(1 + _D * x1 * x2 * y1 * y2)
    y3 = (y1 * y2 + x1 * x2) * _inv(1 - _D * x1 * x2 * y1 * y2)
    return (x3 % _Q, y3 % _Q)


def _scalarmult(p, e: int):
    q = _IDENT
    while e > 0:
        if e & 1:
            q = _edwards(q, p)
        p = _edwards(p, p)
        e >>= 1
    return q


def _encode_point(p: tuple[int, int]) -> bytes:
    x, y = p
    buf = bytearray(32)
    for i in range(31):
        buf[i] = (y >> (8 * i)) & 0xFF
    buf[31] = ((y >> (8 * 31)) & 0x7F) | ((x & 1) << 7)
    return bytes(buf)


def _decode_point(s: bytes) -> tuple[int, int]:
    if len(s) != 32:
        raise ValueError("point length")
    y = sum(s[i] << (8 * i) for i in range(31)) + ((s[31] & 0x7F) << (8 * 31))
    if y >= _Q:
        raise ValueError("y out of range")
    x = _xrecover(y)
    if x & 1 != (s[31] >> 7) & 1:
        x = _Q - x
    p = (x, y)
    if not _on_curve(p):
        raise ValueError("not on curve")
    return p


def _hint(m: bytes) -> int:
    return int.from_bytes(hashlib.sha512(m).digest(), "little")


def _secret_expand(seed: bytes) -> int:
    h = hashlib.sha512(seed).digest()
    a = bytearray(h[:32])
    a[0] &= 248
    a[31] &= 63
    a[31] |= 64
    return int.from_bytes(a, "little")


def ed25519_keygen(seed: bytes | None = None) -> tuple[str, str]:
    """返回 (私钥种子 hex, 公钥 hex)。"""
    if seed is None:
        seed = os.urandom(32)
    pub = _encode_point(_scalarmult(_B, _secret_expand(seed)))
    return seed.hex(), pub.hex()


def ed25519_sign(seed_hex: str, msg: bytes) -> str:
    seed = bytes.fromhex(seed_hex)
    a = _secret_expand(seed)
    h = hashlib.sha512(seed).digest()
    pub = _encode_point(_scalarmult(_B, a))
    r = _hint(h[32:] + msg) % _L
    rp = _encode_point(_scalarmult(_B, r))
    k = _hint(rp + pub + msg) % _L
    s = (r + k * a) % _L
    return (rp + s.to_bytes(32, "little")).hex()


def ed25519_verify(pub_hex: str, msg: bytes, sig_hex: str) -> bool:
    try:
        pub = bytes.fromhex(pub_hex)
        sig = bytes.fromhex(sig_hex)
    except ValueError:
        return False
    if len(pub) != 32 or len(sig) != 64:
        return False
    try:
        a_point = _decode_point(pub)
        r_point = _decode_point(sig[:32])
    except ValueError:
        return False
    s = int.from_bytes(sig[32:], "little")
    if s >= _L:
        return False
    # P2-4：拒绝小阶公钥（[8]A == identity）。本实现用 strict equality 比较
    # （无 cofactor），在无 torsion 分量的诚实公钥上等价于 RFC 8032 的 cofactored
    # 校验；但含小阶分量的公钥可让伪造签名通过，故先剔除（与 libsodium 行为一致）。
    if _scalarmult(a_point, 8) == _IDENT:
        return False
    h = _hint(sig[:32] + pub + msg) % _L  # RFC 8032：k = H(R ‖ A ‖ M)，必须含公钥
    s_b = _scalarmult(_B, s)
    h_a = _scalarmult(a_point, h)
    r_check = _edwards(s_b, ((_Q - h_a[0]) % _Q, h_a[1]))  # sB − hA：负点是 (Q−x) mod Q
    return _encode_point(r_point) == _encode_point(r_check)


# ---------------------------------------------------------------------------
# license 逻辑
# ---------------------------------------------------------------------------
class LicenseError(Exception):
    def __init__(self, code: int, message: str):
        super().__init__(message)
        self.code = code


@dataclass
class LicenseState:
    feature: str
    mode: str  # "free" | "licensed"
    email: str = ""
    expires: str = ""
    features: list = field(default_factory=list)
    source: str = ""


_ENV_PUBKEY_WARNED = False
_DEV_WARNED = False


def _dev_mode() -> bool:
    """开发模式：PDF2PPTX_OMML_DEV=1。仅此模式允许环境变量注入公钥。"""
    return os.environ.get(DEV_MODE_ENV, "").strip() in ("1", "true", "True", "yes")


def _active_pubkey() -> str:
    """发布构建只认模块内常量；环境变量公钥仅在开发模式生效（否则忽略并告警）。"""
    global _ENV_PUBKEY_WARNED, _DEV_WARNED
    env = os.environ.get(PUBLIC_KEY_ENV, "").strip()
    if not env:
        return PUBLIC_KEY_HEX
    if _dev_mode():
        if not _DEV_WARNED:
            _DEV_WARNED = True
            print(
                "WARNING: DEV MODE — 公钥来自环境变量，许可校验不可信（仅开发用）",
                file=sys.stderr,
            )
        return env
    if not _ENV_PUBKEY_WARNED:
        _ENV_PUBKEY_WARNED = True
        print(
            "warning: 已忽略环境变量公钥（非开发模式只信任模块内 PUBLIC_KEY_HEX）",
            file=sys.stderr,
        )
    return PUBLIC_KEY_HEX


def _pubkey_configured() -> bool:
    pk = _active_pubkey()
    if pk == PLACEHOLDER_PUBKEY:
        return False
    try:
        bytes.fromhex(pk)
        return len(pk) == 64
    except ValueError:
        return False


def _find_license_file() -> Path | None:
    env = os.environ.get(LICENSE_ENV, "").strip()
    if env:
        p = Path(env)
        return p if p.is_file() else None
    for cand in (Path("license.json"), Path.home() / ".pdf2pptx_omml" / "license.json"):
        if cand.is_file():
            return cand
    return None


def _canonical_payload(data: dict) -> bytes:
    payload = {k: v for k, v in data.items() if k != "signature"}
    return json.dumps(payload, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")


def require_license(feature: str = "core") -> LicenseState:
    """所有 CLI 入口的统一闸门。

    决策顺序（I2 / P1-2）：
    1. FREE_FEATURES 里的 feature → 永远免费放行（probe 等）；
    2. 没有 license 文件 → 付费 feature（PAID_FEATURES）拒绝（退出码 1），其余免费放行；
    3. 有 license 但本构建未配置公钥 → 告警后按第 2 条规则处理；
    4. 其余：严格验签（签名 → 产品 → 有效期 → 功能范围）。

    注意：PAID_FEATURES 现在为空，所以第 2 条不会触发，行为与历史版本一致；
    上线付费层时只改这个集合，不需要动这里的逻辑。
    """
    if feature in FREE_FEATURES:
        return LicenseState(feature=feature, mode="free", source="always-free")

    lic_file = _find_license_file()
    if lic_file is None:
        if feature in PAID_FEATURES:
            raise LicenseError(1, f"功能 {feature!r} 需要 license（未找到 license 文件）")
        return LicenseState(feature=feature, mode="free", source="no-license-file")

    if not _pubkey_configured():
        print(
            "warning: 检测到 license 文件但本构建未配置公钥，许可体系未启用，按免费模式运行",
            file=sys.stderr,
        )
        if feature in PAID_FEATURES:
            raise LicenseError(1, f"功能 {feature!r} 需要 license（本构建未启用许可体系）")
        return LicenseState(feature=feature, mode="free", source="license-ignored")

    # 严格验证路径（付费层启用后生效）
    try:
        data = json.loads(Path(lic_file).read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("root must be object")
    except (ValueError, OSError) as e:
        raise LicenseError(5, f"license 格式损坏：{e}") from None

    sig = data.get("signature", "")
    if not isinstance(sig, str) or not sig:
        raise LicenseError(5, "license 缺少 signature 字段")
    if not ed25519_verify(_active_pubkey(), _canonical_payload(data), sig):
        raise LicenseError(2, "license 签名无效——文件被修改或与当前发行方不符")

    if data.get("product") != PRODUCT:
        raise LicenseError(4, f"产品不符：期望 {PRODUCT}，实际 {data.get('product')!r}")

    expires = data.get("expires", "")
    if expires:
        try:
            exp = date.fromisoformat(expires)
        except (ValueError, TypeError):
            raise LicenseError(5, f"expires 日期格式非法：{expires!r}") from None
        if exp < date.today():
            raise LicenseError(3, f"license 已于 {expires} 过期")

    feats = data.get("features", [])
    if feature not in feats:
        raise LicenseError(4, f"license 不包含功能 {feature!r}（现有：{feats}）")

    return LicenseState(
        feature=feature,
        mode="licensed",
        email=data.get("issued_to", ""),
        expires=expires,
        features=feats,
        source=str(lic_file),
    )


def require_license_or_exit(feature: str = "core") -> LicenseState:
    """CLI 入口专用：验证失败时打印友好错误并以对应退出码退出（而非抛异常）。"""
    try:
        return require_license(feature)
    except LicenseError as e:
        print(f"error（退出码 {e.code}）：{e}", file=sys.stderr)
        sys.exit(e.code)


# ---------------------------------------------------------------------------
# CLI：密钥生成 / 签发 / 校验 / 自检
# ---------------------------------------------------------------------------
def _cmd_generate_keys(out_dir: str) -> int:
    d = Path(out_dir)
    d.mkdir(parents=True, exist_ok=True)
    seed_hex, pub_hex = ed25519_keygen()
    priv = d / "ed25519_seed.hex"
    pub = d / "ed25519_public.hex"
    priv.write_text(seed_hex + "\n", encoding="utf-8")
    pub.write_text(pub_hex + "\n", encoding="utf-8")
    print(f"私钥种子：{priv}")
    print(f"公钥：    {pub}")
    print()
    print("下一步（上线付费层时）：")
    print(f"  1. 把 {pub} 的内容替换 license_check.py 里的 PUBLIC_KEY_HEX（不要把私钥放进仓库！）")
    print(f"  2. 私钥只在签发时使用：license_check.py --issue --private-key {priv.name} ...")
    return 0


def _cmd_issue(args) -> int:
    seed_path = Path(args.private_key)
    if not seed_path.is_file():
        print(f"error: 私钥文件不存在 {seed_path}", file=sys.stderr)
        return 7
    seed_hex = seed_path.read_text(encoding="utf-8").strip()
    try:
        bytes.fromhex(seed_hex)
    except ValueError:
        print("error: 私钥文件不是合法 hex", file=sys.stderr)
        return 7

    feats = [f.strip() for f in args.features.split(",") if f.strip()]
    issued = date.today()
    expires = issued + timedelta(days=args.days)
    data = {
        "product": PRODUCT,
        "features": feats,
        "issued_to": args.email,
        "issued_at": issued.isoformat(),
        "expires": expires.isoformat(),
    }
    data["signature"] = ed25519_sign(seed_hex, _canonical_payload(data))
    out = Path(args.out) if args.out else Path("license.json")
    out.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"已签发 {out}（{args.email}，features={feats}，至 {expires}）")
    return 0


def _cmd_verify(args) -> int:
    """退出码（P2-3，与模块 docstring 一致）：

    0 = 通过（免费模式，或有效 license；用 status: 行区分 FREE / OK）
    1 = 未找到 license 文件（免费模式，不是错误，但脚本可据此判断"有没有 license"）
    2..7 = 具体失败原因
    """
    lic_arg = getattr(args, "verify", None)  # --verify [路径]（空串/None = 自动查找）
    lic_path = Path(lic_arg) if lic_arg else _find_license_file()
    if lic_path is None or not lic_path.is_file():
        print("status: FREE（未找到 license 文件——免费模式，exit 1）")
        return 1
    if not _pubkey_configured():
        print("status: FREE（本构建未配置公钥，许可体系未启用，exit 0）")
        return 0
    try:
        st = require_license(args.feature)
    except LicenseError as e:
        print(f"status: DENIED（code={e.code}）：{e}")
        return e.code
    print(f"status: OK（mode={st.mode}，{st.email}，features={st.features}，expires={st.expires}）")
    return 0


def _cmd_self_test() -> int:
    """生成→签发→校验 全链路回归，不污染工作目录（全部写临时目录）。

    覆盖：RFC 8032 TEST 1（公钥 + 签名向量）、签发/验签、篡改、过期、越权、
    免费模式、环境变量公钥在非开发模式被忽略、PAID_FEATURES 拒绝、小阶公钥拒绝。
    """
    saved = {k: os.environ.get(k) for k in (PUBLIC_KEY_ENV, LICENSE_ENV, DEV_MODE_ENV)}
    try:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            # --- RFC 8032 §7.1 TEST 1（空消息）：公钥与签名都要逐字节对上 ---
            kat_seed = "9d61b19deffd5a60ba844af492ec2cc44449c5697b326919703bac031cae7f60"
            kat_pub = "d75a980182b10ab7d54bfed3c964073a0ee172f3daa62325af021a68f707511a"
            kat_sig = (
                "e5564300c360ac729086e2cc806e828a84877f1eb8e5d974d873e06522490155"
                "5fb8821590a33bacc61e39701cf9b46bd25bf5f0595bbe24655141438e7a100b"
            )
            gen_pub = _encode_point(_scalarmult(_B, _secret_expand(bytes.fromhex(kat_seed)))).hex()
            assert gen_pub == kat_pub, "RFC 8032 公钥向量失败"
            assert ed25519_sign(kat_seed, b"") == kat_sig, "RFC 8032 签名向量失败"
            assert ed25519_verify(kat_pub, b"", kat_sig), "RFC 8032 验签向量失败"
            assert not ed25519_verify(kat_pub, b"x", kat_sig), "错误消息未被拒"

            # --- 小阶公钥必须拒绝（P2-4），即便签名结构自洽 ---
            small_order_pub = "01" + "00" * 31  # identity 的编码
            assert not ed25519_verify(small_order_pub, b"", kat_sig), "小阶公钥未被拒"

            # --- 非开发模式：环境变量公钥必须被忽略（P1-1） ---
            seed_hex, pub_hex = ed25519_keygen()
            os.environ[PUBLIC_KEY_ENV] = pub_hex
            os.environ.pop(DEV_MODE_ENV, None)
            assert _active_pubkey() == PUBLIC_KEY_HEX, "非开发模式仍采用了环境变量公钥"

            # --- 开发模式：允许环境变量公钥，走完整验签链路 ---
            os.environ[DEV_MODE_ENV] = "1"
            os.environ[LICENSE_ENV] = str(td / "license.json")
            assert _active_pubkey() == pub_hex

            lic = td / "license.json"
            data = {
                "product": PRODUCT,
                "features": ["probe", "core"],
                "issued_to": "selftest@local",
                "issued_at": date.today().isoformat(),
                "expires": (date.today() + timedelta(days=1)).isoformat(),
            }
            data["signature"] = ed25519_sign(seed_hex, _canonical_payload(data))
            lic.write_text(json.dumps(data), encoding="utf-8")
            assert require_license("core").mode == "licensed"

            # 篡改任一字段必须被抓住
            bad = dict(data, issued_to="attacker@local")
            lic.write_text(json.dumps(bad), encoding="utf-8")
            try:
                require_license("core")
                raise AssertionError("篡改未检出")
            except LicenseError as e:
                assert e.code == 2, f"期望 code=2，实际 {e.code}"

            # 过期 license
            expired = dict(data, expires="2020-01-01")
            expired["signature"] = ed25519_sign(seed_hex, _canonical_payload(expired))
            lic.write_text(json.dumps(expired), encoding="utf-8")
            try:
                require_license("core")
                raise AssertionError("过期未检出")
            except LicenseError as e:
                assert e.code == 3, f"期望 code=3，实际 {e.code}"

            # 未付费 feature（license 存在时）
            lic.write_text(json.dumps(data), encoding="utf-8")
            try:
                require_license("pipeline")
                raise AssertionError("越权未检出")
            except LicenseError as e:
                assert e.code == 4, f"期望 code=4，实际 {e.code}"

            # --- 无 license：免费 feature 放行，PAID_FEATURES 拒绝（P1-2） ---
            os.environ.pop(LICENSE_ENV, None)
            os.environ.pop(DEV_MODE_ENV, None)
            os.environ.pop(PUBLIC_KEY_ENV, None)
            assert require_license("probe").mode == "free"
            assert require_license("core").mode == "free"
            old_paid = globals()["PAID_FEATURES"]
            try:
                globals()["PAID_FEATURES"] = frozenset({"pipeline"})
                try:
                    require_license("pipeline")
                    raise AssertionError("付费功能未拦截")
                except LicenseError as e:
                    assert e.code == 1, f"期望 code=1，实际 {e.code}"
                assert require_license("core").mode == "free"  # 非付费 feature 不受影响
            finally:
                globals()["PAID_FEATURES"] = old_paid
    finally:
        for k, v in saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
    print(
        "self-test PASS：RFC 8032 公钥+签名向量 + 签发/篡改/过期/越权 + "
        "免费模式 + 环境变量公钥（非开发忽略 / 开发生效）+ PAID_FEATURES + 小阶公钥拒绝"
    )
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="pdf-to-pptx-omml license 钩子")
    ap.add_argument("--generate-keys", nargs="?", const=".", metavar="DIR", help="生成密钥对到目录（默认当前目录）")
    ap.add_argument("--issue", action="store_true", help="签发 license.json")
    ap.add_argument("--private-key", help="签发用的私钥种子 hex 文件")
    ap.add_argument("--email", default="", help="被授权人（写入 issued_to）")
    ap.add_argument("--features", default="probe,core", help="授权功能列表，逗号分隔")
    ap.add_argument("--days", type=int, default=365, help="有效期天数")
    ap.add_argument("--out", help="签发输出路径（默认 ./license.json）")
    ap.add_argument("--verify", nargs="?", const="", default=None, metavar="LICENSE", help="校验 license（退出码见模块 docstring）")
    ap.add_argument("--feature", default="core", help="--verify / require_license 的功能名")
    ap.add_argument("--self-test", action="store_true", help="全链路自检")
    args = ap.parse_args()

    if args.self_test:
        return _cmd_self_test()
    if args.generate_keys is not None:
        return _cmd_generate_keys(args.generate_keys)
    if args.issue:
        if not args.email:
            print("error: --issue 需要 --email", file=sys.stderr)
            return 7
        return _cmd_issue(args)
    if args.verify is not None:
        return _cmd_verify(args)
    # 无参数： behave as库闸门（供其他脚本 subprocess 调试），打印当前状态
    return _cmd_verify(args)


if __name__ == "__main__":
    sys.exit(main())
