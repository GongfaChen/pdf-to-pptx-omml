#!/usr/bin/env python3
"""license_check.py — 许可证钩子（免费全通过版）。

设计目标（open-core 分层，见 README「License / 授权」一节）：
- 现阶段：无 license 文件 → 一切功能免费全通过（本仓库所有脚本静默放行）。
- 未来付费层：只需 ①生成密钥对并替换下方 PUBLIC_KEY_HEX（或设环境变量），
  ②新付费脚本调用 require_license("pipeline") 等付费 feature。
  老用户环境零破坏，切换当天即生效。

验证语义（退出码与 pdf-image-text 同风格）：
  0 通过（免费模式或有效 license）
  1 未找到 license 文件（仅 --verify 报告时出现，不算错误）
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
纯 Python 实现（点运算走仿射坐标，校验一次约几十毫秒，工具场景足够）。
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
PUBLIC_KEY_ENV = "PDF2PPTX_OMML_PUBKEY"  # 允许用环境变量注入公钥（不重新发版即可换钥）
LICENSE_ENV = "PDF2PPTX_OMML_LICENSE"  # 显式指定 license 文件路径
PRODUCT = "pdf-to-pptx-omml"

FREE_FEATURES = {"probe"}  # 永远免费，无论未来如何配置
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


def _active_pubkey() -> str:
    env = os.environ.get(PUBLIC_KEY_ENV, "").strip()
    return env if env else PUBLIC_KEY_HEX


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
    """所有 CLI 入口的统一闸门。当前版本：免费全通过。

    未来付费层的切换点：
    - probe 类功能永远免费（FREE_FEATURES）。
    - 无 license 文件 → 免费/社区模式放行（保留这一行为即可维持免费层）。
    - 付费 feature + 存在 license + 公钥已配置 → 严格验签，失败按退出码退出。
    """
    if feature in FREE_FEATURES:
        return LicenseState(feature=feature, mode="free", source="always-free")

    lic_file = _find_license_file()
    if lic_file is None:
        return LicenseState(feature=feature, mode="free", source="no-license-file")

    if not _pubkey_configured():
        print(
            "warning: 检测到 license 文件但本构建未配置公钥，许可体系未启用，按免费模式运行",
            file=sys.stderr,
        )
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
    lic_arg = getattr(args, "verify", None)  # --verify [路径]（空串/None = 自动查找）
    lic_path = Path(lic_arg) if lic_arg else _find_license_file()
    if lic_path is None or not lic_path.is_file():
        print("status: FREE（未找到 license 文件——免费模式）")
        return 1
    if not _pubkey_configured():
        print("status: FREE（公钥未配置，许可体系未启用）")
        return 0
    try:
        st = require_license(args.feature)
    except LicenseError as e:
        print(f"status: DENIED（code={e.code}）：{e}")
        return e.code
    print(f"status: OK（mode={st.mode}，{st.email}，features={st.features}，expires={st.expires}）")
    return 0


def _cmd_self_test() -> int:
    """生成→签发→校验 全链路回归，不污染工作目录（全部写临时目录）。"""
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        seed_hex, pub_hex = ed25519_keygen()
        # 已知答案测试：RFC 8032 Test 1（空消息）
        kat_seed = "9d61b19deffd5a60ba844af492ec2cc44449c5697b326919703bac031cae7f60"
        kat_pub = "d75a980182b10ab7d54bfed3c964073a0ee172f3daa62325af021a68f707511a"
        gen_pub = _encode_point(_scalarmult(_B, _secret_expand(bytes.fromhex(kat_seed)))).hex()
        assert gen_pub == kat_pub, "RFC 8032 公钥向量失败"
        os.environ[PUBLIC_KEY_ENV] = pub_hex
        os.environ[LICENSE_ENV] = str(td / "license.json")

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
        bad["signature"] = data["signature"]
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

        # 未付费 feature
        lic.write_text(json.dumps(data), encoding="utf-8")
        try:
            require_license("pipeline")
            raise AssertionError("越权未检出")
        except LicenseError as e:
            assert e.code == 4, f"期望 code=4，实际 {e.code}"

        del os.environ[PUBLIC_KEY_ENV]
        del os.environ[LICENSE_ENV]
        # 无 license → 免费全通过
        st = require_license("pipeline")
        assert st.mode == "free"
        st = require_license("probe")
        assert st.mode == "free"
    print("self-test PASS：RFC 8032 向量 + 签发/篡改/过期/越权/免费 全部符合预期")
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
