# Voucher Codex

银行流水驱动的用友 BIP 记账凭证预生成 demo。

## 项目目标

本项目用于验证一套按日记账的凭证生成流程：

- 每笔银行流水生成一张候选凭证，避免遗漏银行收付。
- 支持财务在预生成凭证界面中修改分录、增加手工凭证、删除非银行锚点分录。
- 每日凭证由用户指定起始凭证号，导出前按顺序连续编号。
- 匹配不到发票或规则不明确的流水，也生成通用待确认凭证。
- 确认无误后再导出导入表，减少在用友 BIP 中二次修改。

## 当前 demo 功能

- 读取银行流水并归一化为统一流水池。
- 支持吉林银行 `.xlsx`、建设银行 `.xls/.xlsx`、工商银行 `.xls/.xlsx` 对账单格式。
- 支持一次导入多个银行流水文件，并显示账户汇总。
- 支持查看全部导入流水或当前日期流水。
- 按日期生成当日候选凭证。
- 支持新增手工凭证。
- 支持编辑凭证日期、凭证号、状态、摘要、科目、借贷金额、客户/供应商/部门/项目/银行档案等字段。
- 支持新增和删除分录；银行流水自动生成的 `1002 银行存款` 锚点分录受保护。
- 支持借贷平衡、缺科目、通用待确认凭证等基础校验。
- 支持导出 CSV 预览。

## 业务规则摘要

### 银行流水

- 凭证日期默认取银行流水交易时间。
- 银行进账：借记 `1002 银行存款`。
- 银行出账：贷记 `1002 银行存款`。
- 每笔银行流水生成一张候选凭证。
- 匹配不到业务规则时生成通用待确认凭证，后续由财务人工修改。

### 发票匹配

- 银行进账尝试按对方户名、金额、日期匹配发票。
- 派遣类发票备注中的 `管理费` 为公司收入，按 6% 计税。
- 派遣类发票备注中的 `扣除额` 为代发工资，不计税，走往来/代收代付。
- `货物或应税劳务名称` 含“外包”的发票按全额计税。

### 每日凭证号

- 按日处理凭证，不跨多日一次性最终编号。
- 每日开始凭证号由用户指定，用于承接用友 BIP 中上一张凭证号。
- 导出前按当前排序从起始号连续编号。

## 目录结构

```text
work/voucher_demo/
  app.py        # 本地 demo 后端，含银行流水解析、凭证生成、校验和导出接口
  index.html    # 单页前端界面

work/*.py       # 数据检查、分析和验证脚本
outputs/        # 生成的分析结果，已被 .gitignore 忽略
work/pydeps/    # 本地安装的运行依赖，已被 .gitignore 忽略
```

## 部署与运行

本 demo 使用 Python 标准库 HTTP 服务，数据读取依赖 `pandas/openpyxl`，老式 `.xls` 支持依赖 `xlrd`。

### 使用 uv 创建虚拟环境

推荐使用 `uv` 管理本地 Python 虚拟环境和依赖。

```bash
uv venv .venv
source .venv/bin/activate
uv pip install pandas openpyxl xlrd
```

启动 demo：

```bash
python work/voucher_demo/app.py
```

然后访问：

```text
http://127.0.0.1:8765
```

### 当前 Codex 环境运行方式

如果在当前 Codex 工作区继续使用已有运行环境，可以这样启动：

```bash
PYTHONPATH=work/pydeps '/Users/sunyitong/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' work/voucher_demo/app.py
```

访问地址同样是：

```text
http://127.0.0.1:8765
```

### 数据路径配置

当前 demo 中样例数据路径写在 `work/voucher_demo/app.py` 顶部：

```python
BANK_PATH = Path('/Users/sunyitong/coding/voucher/代码资料/银行流水/彩虹吉林银行4月对账单.xlsx')
INVOICE_PATH = Path('/Users/sunyitong/coding/voucher/代码资料/4月901张发票.xlsx')
CODE_DIR = Path('/Users/sunyitong/coding/voucher/代码资料/化简代码表')
```

在其他机器部署时，需要把这些路径改成本机实际资料目录。后续正式化时建议改为配置文件或页面上传。

## 数据文件说明

demo 当前使用本机资料目录中的样例数据：

- `/Users/sunyitong/coding/voucher/代码资料/银行流水/彩虹吉林银行4月对账单.xlsx`
- `/Users/sunyitong/coding/voucher/代码资料/4月901张发票.xlsx`
- `/Users/sunyitong/coding/voucher/代码资料/化简代码表/`

导入页面可继续上传工行、建行等银行流水文件。

## 后续计划

- 生成用友 BIP 凭证导入 Excel 模板，而不仅是 CSV 预览。
- 增加持久化数据库，保存每日凭证池、编辑状态和导出记录。
- 增加银行账户映射维护界面。
- 增加客户、供应商、项目、部门、现金流量项目的选择器。
- 增加内部转账识别与合并确认。
- 增加更完整的发票多票合并匹配和红票冲销规则。
- 增加每日确认、锁定、反确认流程。
