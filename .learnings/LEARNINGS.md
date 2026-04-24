# Learnings

## [LRN-20260423-001] stuck_in_planning_loop

**Logged**: 2026-04-23T22:44:00+08:00
**Priority**: high
**Status**: resolved
**Area**: workflow

### Summary
在对大型单文件应用做涉及前后端多处联动的改造时，反复创建 todo list 但不执行代码修改，陷入规划循环。

### Details
用户要求文件管理改造（预览sheet、选择性导入、缓存、关闭sheet），涉及 app.py 后端+前端 HTML 全部改动。连续多个 turn 反复创建/合并 todo list（8个→2个→5个），但一次都没开始写代码。原因是该文件900+行，前端嵌在 Python 字符串中，不敢用 replace_in_file 做多处改动怕出错。

### Suggested Action
对于大型联动改造，应该直接用 write_to_file 整文件重写，而不是反复规划后逐一 replace_in_file。宁可一次写完，不要分步犹豫。

### Metadata
- Source: user_feedback
- Tags: workflow, planning_loop, single_file_app
- Pattern-Key: workflow.full_rewrite_over_incremental_edits
- Recurrence-Count: 1
- First-Seen: 2026-04-23
