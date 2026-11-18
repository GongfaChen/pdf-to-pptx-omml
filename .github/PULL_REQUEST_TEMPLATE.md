## 这个 PR 做了什么

<!-- 一两句话说清改动内容和动机。修 bug 请附 issue 编号。 -->

## 类型

- [ ] 修 bug
- [ ] 新功能
- [ ] 重构（不改变行为）
- [ ] 文档 / 参考资料
- [ ] 新增坑记录（`references/pitfalls.md`）

## 验证方式

<!-- 这个仓库的三道闸是硬要求，请勾选实际跑过的项。 -->

- [ ] `python scripts/omml_helpers.py` —— API 层自检通过
- [ ] `python scripts/license_check.py --self-test` —— 许可证自检通过（未改动该文件可跳过）
- [ ] `python scripts/validate_omml.py lint scripts/*.py` —— 源码 lint 无新增告警
- [ ] `python demo/smoke_test.py demo` —— 构建成功、静态校验 0 issue
- [ ] 在**装有 PowerPoint 的 Windows** 上跑过 smoke test，并确认输出里有 `PowerPoint COM open: PASS` 和 `公式渲染: PASS`
      （Linux/macOS 上这一步会打印 SKIP——**SKIP 不等于通过**，不能作为渲染验收依据）

## 如果是修公式相关的问题

- [ ] 贴了 before / after 的渲染截图（PowerPoint 里截，或 COM 导出 PDF 后截图）
- [ ] 说明这个问题是否已被 `references/pitfalls.md` 里的某条坑覆盖；若是**新坑**，已追加编号记录

## 破坏性变更

<!-- 有无破坏已有构建脚本的接口变化？没有就写「无」。 -->

无

## 其他说明

<!-- 设计上的取舍、已知局限、后续想做的事。 -->
