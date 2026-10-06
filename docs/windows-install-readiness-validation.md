# Windows 安装就绪检查（2026-10-06）

## 修改范围

仅修改表格版。安装界面在复制文件、停止旧服务之前，通过 WMI 检查 wps.exe、et.exe、wpp.exe；有进程或查询失败则停止，不强制结束 WPS。底层 --install 使用 tasklist CSV 再检查一次，注册前再次检查，防止在服务启动期间重新打开 WPS。

安装顺序调整为服务启动 → 健康检查 → 注册 publish.xml → 自启动项。保留失败回滚；只有注册返回成功后，后续失败才允许回滚 publish.xml。注册失败时不清理旧注册，仍停止本次服务。authaddin.json 不读写，不强制改变用户禁用状态。

## 已验证

- npm test：162 项通过。
- node scripts/build-windows-server.js --test：原生 Go 测试通过；新增进程 CSV 检测、未知状态拒绝、安装顺序、四阶段失败回滚及回滚错误测试。
- Windows x64 服务编译通过。
- Inno Setup 安装器编译通过。使用现有 0.2.0 版本参数仅做本地编译检查，未发布或覆盖 GitHub Release。
- git diff --check 通过。
- 本机实际 WMI 查询成功，检测到 7 个 WPS 进程。只执行查询，没有运行安装器、关闭 WPS 或修改已安装环境。
- 文字版工作区无修改。

## 尚未实测

以下验收均未执行，不能据编译成功宣称原 enable:false 问题已解决：

1. WPS 正在运行时执行安装：提示保存并完全退出；旧服务、注册文件、安装文件不被更改。
2. 保存文档、完全退出 WPS（含托盘），执行安装；服务健康后注册成功，重开表格看到功能区。
3. 备份 publish.xml/authaddin.json 后记录安装前后本插件启用状态；若仍出现 enable:false，保留日志继续调查，不自动写 true。
4. 保留文字版安装并验证两套插件均能校对，其他加载项注册不变。
5. 在隔离环境模拟服务启动/健康检查失败，确认旧注册不变；注册或自启动失败时确认本次服务停止、注册按原回滚逻辑恢复。

进程检查不是跨进程锁，无法禁止用户在最终检查之后立即启动 WPS。安装期间仍应保持 WPS 完全退出。此改动消除已知安装时序风险，不证明 WPS 自动禁用的根因。

## 审查回归修复

审查发现原 registrationAttempted 在 register(true) 返回前置为 true，备份或最终替换失败时可能误删原本可用的注册。已将完成状态交由 installSteps 控制：启动、健康检查、注册失败均传 false；仅注册成功后的自启动失败传 true。生产回滚仅在 registrationCompleted 为 true 时处理 publish.xml。

修复分支 fix/windows-install-readiness-v2 基于最新 origin/main（67e59a9，v0.2.2），移植旧安装器提交后修复，保留原远程分支历史。

- 最新基线 npm test：163 项通过。
- 原生 Go 测试通过。新增两种真实文件失败：备份目标被目录占用、Windows 打开原文件阻止最终替换。临时 APPDATA 内预置本插件、文字版和其他插件；实际调用 register(true)，确认旧 publish.xml 字节完全不变。服务和 Run 阶段使用测试替身，确认停止回调被调用、Run 写入阶段未调用且模拟旧 Run 值不变；不触及真实注册表或运行服务。
- 保留注册成功、自启动失败时完成状态为 true 的测试，确保后续失败仍会进入注册回滚。
- Windows 服务和 Inno 安装器重新编译通过；仅本地验证，未安装、未发布。
- git diff --check 通过。

当前真机安装与功能区显示仍未实测；本地验证不能替代远程 CI 结果。
