### 示例

- 目录表格

|目录等级|目录名称|目录说明|
|---|---|---|
|根目录|共享文件|共享文件根目录|
|一级|部门文件|各部门文件|
|二级|信息部|信息部共享文件|
|三级|培训文件|各系统培训用文件|
|二级|销售部|销售部共享文件|
|一级|公司流程|公司各流程目录|

- 目录结构
```
共享文件
  ├部门文件
  |  ├信息部
  |  | └培训文件
  |  └销售部
  └公司流程
```

依据这样的表格信息，在每一行后，拼接成一个完成的共享文档路径，例如：
\\192.168.1.1\共享文件\部门文件\信息部\培训文件