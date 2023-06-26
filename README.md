# CVE_Check_Tool

这是一个用于查询组件CVE漏洞信息的Python脚本。它使用NIST NVD的RESTful API来获取CVE数据，并提供格式化的输出和导出到Excel功能。

## 依赖

- Python 3.x

- 以下Python库：

  - argparse

  - requests

  - prettytable

  - pandas

  - openpyxl

## 安装依赖

```Shell
pip install -r requirements.txt
```




## 使用方法

运行脚本并提供以下参数：

```Python
python CVE_Check_Tool.py -p <产品名称> -v <产品版本> [-o <输出文件>]
```




- `-p` 或 `--product`：要查询的组件名称。

- `-v` 或 `--version`：要查询的组件版本。

- `-o` 或 `--output`（可选）：指定导出到Excel文件的文件名。

如果未提供输出文件名，将默认使用 `<产品名称>_<产品版本>_CVE.xlsx` 作为文件名。

查询结果将显示在控制台，并可选择导出到Excel文件。



## 示例

查询OpenSSL 1.0.1g版本的CVE信息，并将结果导出到Excel文件：



```Python
 python CVE_Check_Tool.py -p nginx -v 0.6。4  -o cve_info.xlsx
```

![image](https://github.com/LIHAQI/CVE_Check_Tool/assets/57976650/56b1250d-a56e-4292-bfad-620edb6bedd1)



## 声明

本脚本仅用于学习和演示目的，使用时请遵守相关法律法规和服务提供商的条款和条件。

