# icost-bookkeeping

把微信 / 支付宝账单变成 [iCost](https://apps.apple.com/app/id1486046331) 能直接导入的文件 —— 一个自包含的记账技能包：脚本、配置模板、官方导入模板骨架、完整文档全在里面，解压即用。

**不依赖任何外部项目目录，也不依赖 WorkBuddy**：你可以把它装成 AI 助手的技能，也可以纯手工跑脚本。

## 它能做什么

- **解析账单**：读取微信支付 xlsx / 支付宝 csv，自动识别列结构
- **五层自动归类**：语义规则 → 商户映射表 → 历史学习 → 品类词 → 平台兜底，实测 91% 自动完成
- **跨账单去重**：以「日期 + 金额绝对值」为键剔除重复记录；同日多笔小额自动合并为群收款（AA）
- **按账本拆分**：iCost 导入文件**一个文件只能含一个账本**，自动拆成多份
- **生成官方兼容的导入文件**：克隆 iCost 官方模板的 XML 结构生成 xlsx（不是用 openpyxl 直接 save —— iOS 解析器对格式极其挑食）
- **退款 / 报销单独处理**：这几类 iCost 导入不识别，自动转成「手工补录清单」
- **逐字段校验**：导入前跑一遍校验报告，把问题拦在导入前

## 安装

### 作为 AI 助手技能（WorkBuddy，推荐）

```bash
git clone https://github.com/coolbarry/icost-bookkeeping.git ~/.workbuddy/skills/icost-bookkeeping
```

WorkBuddy 的 skill 就是一个含 `SKILL.md` 的文件夹，clone 到位即生效，**不需要上传任何市场**：

| 安装位置 | 作用范围 |
|---|---|
| `~/.workbuddy/skills/icost-bookkeeping/` | 用户级，所有项目可用 |
| `<某个项目>/.workbuddy/skills/icost-bookkeeping/` | 项目级，仅该项目可用 |

装好后直接对 AI 说「记账」就会触发。

升级（该目录本身是 git 仓库）：

```bash
git -C ~/.workbuddy/skills/icost-bookkeeping pull
```

不想用 git 也可以：点仓库页面的 **Code → Download ZIP**，解压后把 `icost-bookkeeping-main` 目录放进上面两个位置之一即可。

### 不用 git、也不用 AI

脚本不依赖任何 AI，只需要 Python 3 + openpyxl。解压后直接按下面「快速开始 → 方式二」跑即可。

## 快速开始

### 方式一：作为 AI 助手技能使用（推荐）

先按上面「安装」把仓库 clone 到 `~/.workbuddy/skills/`。之后直接说「记账」，AI 会按 `SKILL.md` 的流程执行：先铺工作目录 → 体检配置 → 解析账单 → 你在确认页里核对少数待定笔 → 生成导入文件。

### 方式二：纯命令行使用

```bash
# 1. 铺一个工作目录（个人配置和数据都放这里）
./scripts/init_project.sh ~/我的记账

# 2. 配置体检（退出码 2=未配置，1=有错，0=就绪）
cd ~/我的记账/icost-system && ./check.sh

# 3. 按 references/初次引导.md 完成配置（需要 iCost 的分类/账户截图）
#    跑绿后删掉 config/SETUP_PENDING.md

# 4. 每次记账
./run.sh   ../账单/微信账单.xlsx ../账单/支付宝账单.csv ../账单/iCost历史导出.xlsx
#   → 在 output/记账确认.html 里确认待定笔数，复制结果回来
./split.sh ../已确认明细.xlsx ../导入文件
```

唯一依赖：**Python 3 + openpyxl**（`lib.sh` 会自动探测可用的 Python）。

## 目录结构

```
icost-bookkeeping/
├── SKILL.md                技能入口：完整流程 + 三条红线 + 故障快查
├── scripts/
│   ├── init_project.sh     铺 / 升级工作目录（--upgrade 保留个人配置）
│   ├── lib.sh              Python 探测等公共函数
│   ├── run.sh              账单解析 → 归类 → 去重 → 确认页
│   ├── split.sh            按账本拆分 → 生成导入文件 + 校验报告
│   ├── check.sh            配置体检
│   └── core/
│       ├── engine.py       归类引擎（五层策略 + 去重 + AA 合并）
│       ├── export_split.py 导入文件生成（克隆官方模板 XML）
│       ├── make_confirm.py 确认页生成
│       └── check_config.py 配置校验
├── references/             完整文档（7 篇）
│   ├── 初次引导.md          第一次使用：6 个问题 + 3 样材料 + 诊断
│   ├── 给AI的指令.md        纯文本 prompt，跨 AI 通用
│   ├── 01_工作流SOP.md      标准作业流程
│   ├── 02_导入规范.md       10 列导入格式逐字段说明
│   ├── 03_踩坑记录.md       已知坑与解法
│   ├── 04_配置与维护.md     config 各文件说明
│   └── 05_复盘方法论.md     基于账本数据的复盘方法
└── assets/
    ├── config/             脱敏配置模板 + 官方导入模板骨架（别删 icost_template.xlsx）
    └── example/            虚构样例，可立即跑通验证
```

## 三条红线（违反必定导入失败）

1. **导出格式 ≠ 导入格式。** iCost 导出是 21 列（文件名叫「仅供查看，不用于数据恢复」），导入只有 **10 列**。`run.sh` 的中间产物不能直接导，必须过 `split.sh`。
2. **一个文件只能含一个账本。** 导入文件没有账本列，账本是在 App 里切换后导入的。
3. **导入只认 `支出 / 收入 / 转账`。** `退款入账 / 报销入账 / 待报销 / 已报销` 写了必然解析失败，这几笔走「手工补录清单」。

## 配置说明

最小可用配置只要 3 个文件，都在 `assets/config/` 有模板：

| 文件 | 作用 |
|---|---|
| `categories.json` | 分类树（账本 › 收支方向 › 一级 › 二级），**必须从 App 分类管理截图录入**，不能从历史数据反推 |
| `accounts.json` | 账户体系 |
| `pay_mapping.json` | 微信/支付宝「支付方式」字段 → 账户映射 |

可选的 `rules.json` / `merchant_map.json` / `import_fix.json` 用得越久越省事：归类错的商家补进 `merchant_map.json`，人工确认的笔数会持续下降。

> ⚠️ `config/` 是个人数据（账户余额、银行尾号、真实商户名），**只写进工作目录，不要写进技能包**。

## 故障快查

| 现象 | 处理 |
|---|---|
| 导入「解析失败」 | 先试导入包内 `00_诊断_官方原版模板.xlsx`；连原版都失败，说明是导入方式不对（走 iOS「文件」App 分享进 iCost），改文件无用 |
| `ModuleNotFoundError: openpyxl` | `pip install openpyxl` |
| 自动归类率 < 60% | 多半是没传历史账单导出，或分类树漏了账本 |
| 账单读不出来 | 平台格式变了，改 `core/engine.py` 的列索引 |

## 设计说明

导入文件不用 openpyxl 直接 save，而是**克隆 iCost 官方模板的 XML 结构**（`sheet1.xml` + `sharedStrings.xml`）再替换内容 —— 因为 iCost 的 iOS 解析器对富文本、inlineStr、数字实体等写法极其挑食，只有官方同款结构能稳定导入。踩坑全过程见 `references/03_踩坑记录.md`。

## License

MIT
