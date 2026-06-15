# 脚本说明

当前目录提供一个可直接扩展的教师详情补采脚本：

- [teacher_profile_pipeline.js](/Users/xinxinhuashe/Documents/易达威实习/scripts/teacher_profile_pipeline.js)
- [统计有效数据.py](/Users/xinxinhuashe/Documents/易达威实习/scripts/统计有效数据.py)

## 作用

把八爪鱼导出的教师列表 CSV 进一步整理成导入模版格式，默认按“学校文件夹 / 学院文件”落盘为 `.csv`，同时生成一个总的 `按院系汇总.xlsx` 方便统一上传。

推荐输入字段：

- `name` 或 `姓名`
- `profile_url` 或 `主页`
- `school` 或 `学校`
- `college` 或 `学院`

脚本会尽量自动补：

- `职称`
- `邮箱`
- `简介`

## 最小使用流程

1. 把八爪鱼导出的 CSV 放到 `原始数据/八爪鱼导出/`
2. 执行脚本
3. 在 `输出结果/学校名/学院名.csv` 查看单个院系结果
4. 在 `输出结果/日期/按院系汇总.xlsx` 查看统一上传文件
5. 在 `日志/` 查看失败和缺失项
6. 同一天的数据会自动生成：

- `按院系汇总.xlsx`

## 运行示例

```bash
node /Users/xinxinhuashe/Documents/易达威实习/scripts/teacher_profile_pipeline.js \
  --input "/Users/xinxinhuashe/Documents/易达威实习/原始数据/八爪鱼导出/示例列表.csv" \
  --output "/Users/xinxinhuashe/Documents/易达威实习/输出结果/示例输出.csv" \
  --collector "你的名字" \
  --date "2026-06-01" \
  --concurrency 5
```

## 单页列表型师资页

对于“一整个页面里直接列出多位老师信息”的页面，可以直接用主脚本的新模式生成标准 CSV，并汇总到同一天的 `按院系汇总.xlsx`：

```bash
python3 /Users/xinxinhuashe/Documents/易达威实习/scripts/teacher_profile_pipeline.py \
  --list-page-url "https://sast.xidian.edu.cn/szdw/ggjs1/kjkxyjssyzx.htm" \
  --output "/Users/xinxinhuashe/Documents/易达威实习/输出结果/2026-06-04/西安电子科技大学/空间科学与技术实验中心-西安电子科技大学空间科学与技术学院输出.csv" \
  --school "西安电子科技大学" \
  --college "空间科学与技术学院" \
  --collector "你的名字" \
  --date "2026-06-04"
```

## 单个教师主页

如果你手上是某位老师的个人主页 URL，也可以直接补采这一位老师：

```bash
python3 /Users/xinxinhuashe/Documents/易达威实习/scripts/teacher_profile_pipeline.py \
  --profile-url "https://faculty.xauat.edu.cn/liyanjun/zh_CN/index.htm" \
  --school "西安建筑科技大学" \
  --college "材料科学与工程学院" \
  --collector "你的名字" \
  --date "2026-06-10"
```

## 调试建议

- 第一次跑时先加 `--limit 20`
- 先挑一个网站结构比较规整的学院页面试跑
- 看输出 CSV 和总汇总 XLSX 是否满足导入要求
- 看日志里失败链接是否集中在某一个学校站点

## 后续增强方向

1. 为常见高校站点增加站点规则
2. 对简介字段加入更精细的正文定位
3. 加入断点续跑
4. 支持读取 Excel 而不仅是 CSV
5. 支持从飞书任务池自动读取待处理院系

## 统计当天有效数据

```bash
python3 /Users/xinxinhuashe/Documents/易达威实习/scripts/统计有效数据.py --date "2026-06-04"
```

有效数据规则：

- 除 `序号` 外，导入模板里的字段都非空
