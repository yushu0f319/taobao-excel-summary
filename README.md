# 淘宝日度 Excel 更新 Skill

这个项目用于把淘宝本地导出的三张原始表，追加进固定的日度规模数据 Excel 模板。

当前 v0.1 只处理淘宝三张表：

- `TB门店`
- `TB流量`
- `TB履约`

## 交付形态

这是一个 Codex Skill，不是网页系统，也不需要自动登录淘宝。

用户提供：

1. 上一日成品表，例如 `6月4日度-规模数据.xlsx`
2. 当日淘宝导出的三张原始表

输出：

- 新一日成品表，例如 `6月5日度-规模数据.xlsx`

## 安装

把这个目录复制到 Codex skills 目录：

```bash
skills/taobao-daily-excel-update
```

也就是最终结构类似：

```text
~/.codex/skills/taobao-daily-excel-update/
├── SKILL.md
└── scripts/
    └── update_taobao_daily.py
```

## 使用

在 Codex 里选择或触发 `taobao-daily-excel-update`，然后告诉它：

```text
请用淘宝日度 Excel 更新流程：
底稿是 /path/to/6月4日度-规模数据.xlsx
原始表在 /path/to/taobao-raw/
输出到 /path/to/6月5日度-规模数据.xlsx
```

Skill 会自动：

1. 判断三张原始表分别是 `TB门店`、`TB流量`、`TB履约`
2. 追加到成品表对应 sheet
3. 添加空行、表头和数据
4. 把日期、门店 ID、金额、订单、人数、百分比等转换成模板可识别格式
5. 校验写入值和源数据一致

## 命令行运行

也可以直接运行脚本：

```bash
python3 skills/taobao-daily-excel-update/scripts/update_taobao_daily.py \
  --base "/path/to/6月4日度-规模数据.xlsx" \
  --input "/path/to/taobao-raw-folder" \
  --output "/path/to/6月5日度-规模数据.xlsx"
```

## 测试

```bash
python3 -m unittest tests/test_update_taobao_daily.py -v
```
