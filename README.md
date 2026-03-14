# ChemStation 单次运行触发脚本

这个仓库用于通过 Agilent ChemStation 的 DDE 接口触发一次单次运行（single run）。

## 文件说明

- `start_single_run.mac`：ChemStation 宏文件。会先执行 `PrepRun`，并在检测到 `ACQSTATUS$ = "PRERUN"` 时调用 `StartMethod`。
- `chemstation_start_run.py`：Python 触发脚本。通过 DDE 发送宏执行命令，并轮询状态确认系统从 `READY` 离开。
- `tests/test_start_single_run_mac.py`：针对宏文件的结构与关键命令测试。

## 使用前准备

1. 在 ChemStation 中确认宏文件路径可访问。
2. 根据实际环境修改 `chemstation_start_run.py` 中的配置：
   - `CHEMSTATION_APP`
   - `MACRO_PATH`
   - `TOPIC_EXECUTE`
3. 安装依赖：

```bash
pip install pywin32
```

> 说明：`pywin32` 和 DDE 仅适用于 Windows 环境。

## 运行脚本

```bash
python chemstation_start_run.py
```

## 运行测试

```bash
pytest -q
```

测试主要验证 `start_single_run.mac` 是否包含预期流程：

- 调用 `PrepRun`
- 存在 `while loop = 1`
- 在 `PRERUN` 条件下调用 `StartMethod`
- 以 `EndMacro` 结束
