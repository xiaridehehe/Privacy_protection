# 医疗病历及 DICOM 影像批量隐私脱敏

这是一个离线、独立的 Python 命令行程序。支持 `records`（病历）、`dicom`（医学影像）和 `both`（病历与影像配对处理）。程序默认复制后脱敏，不覆盖输入数据。

> 自动规则和复扫只能降低风险，不能证明数据已经完全匿名化。`manual_review` 中的文件必须经过人工审核和必要的专用工具处理后，才能进入共享数据集。

## 安装（Windows PowerShell）

建议使用 Python 3.10–3.12：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 输入和输出

输入根目录的直接子目录被视为相互独立的患者。即使两个目录同名，也不会因姓名自动合并；稳定映射依据输入根目录内的相对目录键生成，而不是原病案号。

```text
E:\dataset\Medical Record\
  张三\...
  李四\...
```

输出目录结构：

```text
E:\deidentified\
  share_ready\Patient_A8F3C92D\...
  manual_review\Patient_A8F3C92D\...
  deidentification_report.csv
  deidentification.log
```

映射表必须位于输出目录之外的受保护位置。它包含真实姓名、随机 Patient ID 和不可逆的源目录键。程序每分配一个 ID 就原子保存 Excel；重新运行会复用映射。请使用 BitLocker/EFS、受限 NTFS ACL 或机构认可的密钥管理方案保护该文件，并单独备份。

## 运行

病历：

```powershell
python .\medical_deidentifier.py --mode records `
  --input "E:\dataset\Medical Record" `
  --output "E:\dataset_deidentified\records" `
  --mapping "E:\protected\Patient_ID_Mapping.xlsx"
```

DICOM 影像：

```powershell
python .\medical_deidentifier.py --mode dicom `
  --input "E:\dataset\Medical Image" `
  --output "E:\dataset_deidentified\dicom" `
  --mapping "E:\protected\Patient_ID_Mapping.xlsx"
```

病历和影像同时处理并共用 Patient ID：

```powershell
python .\medical_deidentifier.py --mode both `
  --records-input "E:\dataset\Medical Record" `
  --dicom-input "E:\dataset\Medical Image" `
  --output "E:\dataset_deidentified\paired" `
  --mapping "E:\protected\Patient_ID_Mapping.xlsx"
```

`both` 模式使用两个输入根目录中完全相同的患者子目录名称建立明确配对。例如病历目录的 `张三` 只会与影像目录的 `张三` 配对；配对成功时只生成一个 Patient ID，并输出到：

```text
paired\records\share_ready\Patient_A8F3C92D\...
paired\dicom\share_ready\Patient_A8F3C92D\...
```

任一侧缺少对应患者目录时，程序不会终止：病历存在就处理病历，影像存在就处理影像。终端和受控日志会用脱敏后的 Patient ID 提示 `missing=records` 或 `missing=dicom`，不会打印原患者姓名。程序不会根据病历正文姓名或相似名称猜测配对。
如果其中一侧以前已用同一映射表单独处理，`both` 会复用已有 Patient ID，并把另一侧绑定到该 ID；若两侧历史上已被分配不同 ID，程序会拒绝自动合并并要求人工核对映射。

默认发现已有患者输出时安全跳过，适合中断后续跑。确认需要重做已有患者时增加 `--overwrite-patient`。程序返回码：`0` 全部自动处理；`1` 存在人工复核项；`2` 存在错误或启动失败。

## 处理范围和边界

- 病历：支持 TXT/CSV/JSON/XML/DOCX/XLS/XLSX/PDF 等；`.xls` 按旧版 Excel 单元格处理并保留单元格样式，其他格式清理可访问的正文、表格、页眉页脚、批注和文档属性，随后进行写后残留复扫。
- 按当前业务规则，PNG/JPG/TIFF 等普通图像被视为安全文件，内容不做 OCR 或遮盖，文件名脱敏后原样复制到 `share_ready`。老式 DOC、压缩包及无法可靠识别的格式仍进入 `manual_review`。
- 扫描 PDF、PDF 内图片、外部超链接和公式内疑似敏感内容进入人工复核，防止破坏医学内容或漏掉烧录文字。
- DICOM：递归处理序列、删除私有标签、替换身份标签、一致重映射关键 UID、按患者稳定偏移日期、保持像素数据不变，并写后重读验证。
- 影像目录只修改三个目标：`LABEL.TXT`、`DICOMOBJ` 下的 DICOM 文件和 `DICOMDIR`。`DICOMDIR` 必须同步脱敏并重映射引用 UID，否则会破坏其与 DICOMOBJ 的关联。
- `INDEX.files`、`PD-SViewer`、`asdf.txt`、ICO、INF、EXE、CHM、HTML、XML、普通预览图及其他非目标资源保持文件名、目录位置和文件字节完全不变。
- 影像处理采用患者目录级原子规则：目录内 DICOM、LABEL、辅助文本、图像及其他文件必须全部处理成功，整个 Patient ID 才会进入 `share_ready`；任一文件失败或需要复核时，整组移动到 `manual_review`。
- DICOM 仅清理文件内的患者记录标签、私有标签、标识符和日期；像素数据不进行文字遮盖或重编码，并在写后比较原始像素字节。
- 非法组合的 DICOM 字符集声明（如 `GB18030` 同时组合 `ISO_IR 100`）在读取时按 GB18030 解释，脱敏输出统一声明为 UTF-8 (`ISO_IR 192`)。无法完整解码或源 LO 字段超长等源数据规范问题汇总写入报告，不在终端重复刷屏，也不会静默截断非隐私字段。
- 病历按文件独立处理。失败或需要复核时，终端会打印 Patient ID、已经脱敏的文件名和状态，便于定位记录。
- `BurnedInAnnotation` 未明确为 `NO` 的含像素 DICOM、DICOMDIR 和特殊文件进入人工复核。程序不声称能自动识别或遮盖所有像素烧录文字。
- 日志不记录原始姓名、完整手机号、身份证号或原始文件名，只记录 Patient ID 和短哈希。报告的路径使用已脱敏文件名。

本工具采用与 DICOM PS3.15 Basic Application Confidentiality Profile 相符的保守思路，但不是一份经认证的完整合规实现。正式共享前应结合伦理审批、数据使用目的、机构制度及 GB/T 22239-2019、个人信息保护和医疗数据分类分级要求进行人工复核与风险评估。
