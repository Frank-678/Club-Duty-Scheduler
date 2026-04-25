# 社团值班排班器 V4

V4 目标：保留图形化前端，同时加入“每人每周最多排班次数”的全局设置，并自动判断这个上限是否偏大或偏小。

## 时间范围

- 周：`1` 到 `7`
  - `1` 周一
  - `2` 周二
  - `3` 周三
  - `4` 周四
  - `5` 周五
  - `6` 周六
  - `7` 周日
- 值班段：`1` 到 `6`

## 默认开班规则

网页默认使用：

```javascript
function weekdayDefaultOpenShifts() {
  const open = makeOpenShifts(0);
  ['2','3','4','5'].forEach(day => ['1','2','3','4'].forEach(slot => open[day][slot] = 1));
  ['6','7'].forEach(day => ['1','2','3','4'].forEach(slot => open[day][slot] = 1));
  return open;
}
```

也就是默认周二到周日的 1~4 班开班；周一、5~6 班默认不开班。

## 核心功能

- 图形化设置开班/不开班时段
- 图形化新增、删除成员
- 图形化勾选每个成员的课表
- 每个成员设置“优先多排 / 正常 / 优先少排”
- 每个成员设置禁排日期
- 每人每周最多排班次数可调整
- 每个开班时段最多一个人
- 只在没课时间排班
- 能排多少排多少，返回最大可行排班
- 自动提示“最多排班次数”是否偏大或偏小
- 导出 JSON
- 安装 `openpyxl` 后可导出 Excel

## 新增字段

```json
{
  "config": {
    "max_shifts_per_member": 2
  }
}
```

含义：每个成员本周最多被安排几次班。默认是 `1`。

## 排班次数建议逻辑

程序会先尊重用户设置的 `max_shifts_per_member = m` 来排班，同时额外测试不同上限下最多能排多少班。

输出字段：

- `capacity_advice.advice_type`
  - `ok`：当前设置合理
  - `too_large`：m 偏大，更小的 n 已经够用
  - `too_small`：m 偏小，提高到 n 才能排更多或排满
  - `impossible_even_with_more`：即使继续提高上限也排不满，原因是课表/禁排/无人可排
- `capacity_advice.recommended`：建议的每人最多排班次数
- `summary.max_shifts_per_member_requested`：用户设置的 m
- `summary.recommended_max_shifts_per_member`：系统建议值
- `summary.max_shifts_used_in_result`：结果中单个人实际最多排了几次

## 输出语义

- `success: true`：程序正常算出了结果
- `all_filled: true`：所有开班时段都排满
- `all_filled: false`：只能排满一部分
- `status`
  - `full`：全部排满
  - `partial`：部分排满
  - `empty`：一个班都没排上

注意：无法全部排满不是程序失败，而是当前约束下只能得到部分可行解。

## JSON 输入格式

```json
{
  "config": {
    "max_shifts_per_member": 2,
    "open_shifts": {
      "1": {"1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6": 0},
      "2": {"1": 1, "2": 1, "3": 1, "4": 1, "5": 0, "6": 0},
      "3": {"1": 1, "2": 1, "3": 1, "4": 1, "5": 0, "6": 0},
      "4": {"1": 1, "2": 1, "3": 1, "4": 1, "5": 0, "6": 0},
      "5": {"1": 1, "2": 1, "3": 1, "4": 1, "5": 0, "6": 0},
      "6": {"1": 1, "2": 1, "3": 1, "4": 1, "5": 0, "6": 0},
      "7": {"1": 1, "2": 1, "3": 1, "4": 1, "5": 0, "6": 0}
    }
  },
  "members": [
    {
      "name": "张三",
      "priority": "prefer_more",
      "ban_days": ["5"],
      "schedule": {
        "1": {"1": 1, "2": 1, "3": 0, "4": 0, "5": 0, "6": 1},
        "2": {"1": 0, "2": 1, "3": 0, "4": 0, "5": 1, "6": 0},
        "3": {"1": 0, "2": 0, "3": 1, "4": 1, "5": 0, "6": 0},
        "4": {"1": 1, "2": 0, "3": 0, "4": 0, "5": 0, "6": 1},
        "5": {"1": 0, "2": 0, "3": 1, "4": 0, "5": 0, "6": 0},
        "6": {"1": 1, "2": 0, "3": 0, "4": 1, "5": 0, "6": 0},
        "7": {"1": 0, "2": 0, "3": 0, "4": 0, "5": 1, "6": 0}
      }
    }
  ]
}
```

## 命令行使用

```bash
pip install -r requirements.txt
python app.py tests/inputs/case_too_large_full_realistic.json --json result.json --xlsx result.xlsx
```

如果只需要 JSON，不需要 Excel：

```bash
python app.py tests/inputs/case_too_large_full_realistic.json --json result.json
```

## 启动网页表单

```bash
pip install -r requirements.txt
python app.py --web
```

浏览器打开：

```text
http://127.0.0.1:5000
```

## 运行测试

```bash
python run_tests.py
```

测试集包括：

- `case_too_large_full_realistic.json`：用户设置上限偏大，1 次已经足够排满
- `case_too_small_needs_repeat_realistic.json`：用户设置上限偏小，需要提高到 3 次才可排满
- `case_impossible_even_with_more_realistic.json`：即使多排也无法排满，有些时段无人可排
- `case_priority_repeat_realistic.json`：验证优先级和重复排班上限共同生效

## 算法说明

程序使用最小费用最大流：

1. 先尽量排满更多开班时段。
2. 每个人可以有第 1 次、第 2 次、第 3 次……排班容量。
3. 第 n+1 次排班会比第 n 次有更高成本，所以系统会优先让更多人各排一次；只有排不满时才尝试让已排过的人继续排。
4. 同一层次数内，再根据 `prefer_more` / `normal` / `prefer_less` 调整优先级。
5. 最后额外测试不同 `max_shifts_per_member` 下的最大可排数量，生成建议值。

## 备注

旧字段 `require_open_shift_count_less_than_member_count` 已废弃。现在程序不会因为开班数和人数关系直接拒绝，而是总是尝试返回当前约束下的最大可行排班。
