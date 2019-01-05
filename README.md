- [中文](#1)
- [English](#2)

<h2 id = "1"></h2>

## 使用Apache POI 读写Excel，支持.xls和.xlsx格式

### 业务需求
- 业务上需要处理从数据库中导出的Excel，所以选择开源的Apache POI处理

### 业务描述

- [questionnaire_simple](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/questionnaire_simple.xlsx)文件简要记录了原数据和处理后的数据的格式和字段，在此解释需要处理的核心字段，其余字段均可忽略。
- [questionnaire_complete](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/questionnaire_complete.xlsx)文件包含了全部的问卷调查样本。

| field          | meaning            |
| ------------ | ---------------- |
| orderNum |   用户提交的问卷编号 |
| titleKey |   问题序号 |
| titleName|   问题 |
| answer |   答案序号 |

### 业务处理
- 原数据中用户提交问卷的编号对应多条记录，回答每一个问题均对应一条记录，需要将多条记录合并成一条。合并后的表格头由原数据表格基本信息和问题组成。例如：

**原数据**

| orderNum | submitTime | titleKey | titleName | answer | other |
| -------- | ------ | ------ | ------ |------ |------ |
| 79 | 9/26/2018 18:26:32 | 1 | 1.资产规模大致范围? | 3 | …… |
| 79 | 9/26/2018 18:26:32 | 2 |2.理财规模大致范围？| 2 | ……|
| 79 | 9/26/2018 18:26:32 | 3 |3.面临调整压力是？| 4 | ……|

**合并后**

| 序号 | 提交时间 | 1.资产规模大致范围?  | 2.理财规模大致范围？ | 3.面临调整压力是？ | other |
| -------- | -------- | ------ | ------ |------ |------ |
| 1 | 9/26/2018 18:26:32 | 3 | 2| 4 | …… |

其中，第三列，第四列，第五列分别对应问题的答案需要。

- 每个用户回答的问题可能并不是所有的问题，因为问卷可能会根据用户回答的不同进而跳跃到不同的选项中，例如一共10题问卷调查，用户一回答了1，3，5，7，8，10题，但用户二可能回答了1，2，4，5，9，10题，也有可能用户回答了所有问题。

- 问卷可能会有多选。

## 使用
项目提供了 [ExcelPoi.jar](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/ExcelPoi.jar)

### 运行环境
- JRE1.8或JDK1.8或以上

### 运行命令
```java -jar ExcelPoi.jar fileUrl saveUrl saveFileName```
- fileUrl表示原数据路径，例如/home/root/question.xlsx或C:\question.xlsx
- saveUrl表示数据处理后保存的路径，例如/home/root/或C:\
- saveFileName表示数据处理后最终保存的文件名，可选，如果不填写，则默认处理后的文件名为"原文件名_处理后.xlsx"

## API参考
- [How to create a new workbook](http://poi.apache.org/components/spreadsheet/quick-guide.html#NewWorkbook)
- [How to create a sheet](http://poi.apache.org/components/spreadsheet/quick-guide.html#NewSheet)
- [How to create cells](http://poi.apache.org/components/spreadsheet/quick-guide.html#CreateCells)
- [Iterate over rows and cells](http://poi.apache.org/components/spreadsheet/quick-guide.html#Iterator)
- [Getting the cell contents](http://poi.apache.org/components/spreadsheet/quick-guide.html#CellContents)
- [Reading and writing](http://poi.apache.org/components/spreadsheet/quick-guide.html#ReadWriteWorkbook)

## 项目构建
### 使用maven构建
- 添加poi依赖，支持xls格式
```
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->
<dependency>
   <groupId>org.apache.poi</groupId>
   <artifactId>poi</artifactId>
   <version>3.17</version>
</dependency>
```

- 添加poi-ooxml依赖，支持xlsx格式
```
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>
   <groupId>org.apache.poi</groupId>
   <artifactId>poi-ooxml</artifactId>
   <version>3.17</version>
</dependency>
```

- 如果出现```java.lang.NoClassDefFoundError: org/apache/xmlbeans/XmlObject``` 异常，仍需要添加xmlbeans依赖
```
<!-- https://mvnrepository.com/artifact/org.apache.xmlbeans/xmlbeans-->
<dependency>
   <groupId>org.apache.xmlbeans</groupId>
   <artifactId>xmlbeans</artifactId>
   <version>3.0.2</version>
</dependency>
```

### 支持.xls和.xlsx格式
如果原文件是.xls格式，则使用HSSFWorkbook创建Workbook，如果是.xlsx格式，则用XSSFWorkbook创建。

### POI 涉及概念解释
- Workbook：Excel工作簿
- Sheet：工作表
- Row：行
- Cell：单元格

如下图所示：

![工作簿](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/workbook_cn.png)


<h2 id = "2"></h2>

## Read and write Excel files with Apache POI, supporting .xls and .xlsx formats

### Business requirements
- Since we need to process Excel exported from database, we choose open source Apache POI for processing.

### Business description

- [questionnaire_simple](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/questionnaire_simple.xlsx) file briefly records the format and fields of original data and processed data. The core fields that need to be processed are explained here, and the remaining fields can be ignored.
- [questionnaire_complete](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/questionnaire_complete.xlsx) file contains a full sample of questionnaires.

| field          | meaning            |
| ------------ | ---------------- |
| orderNum |   questionnaire number submitted by user |
| titleKey |   question number |
| titleName|   question |
| answer |   answer number |

### Business processing
- In the original data, the number of the questionnaire submitted by user corresponds to multiple records. Each question corresponds to one record, and multiple records need to be merged into one. The merged table header consists of basic data and problems of the original data table. E.g:

**original data**

| orderNum | submitTime | titleKey | titleName | answer | other |
| -------- | ------ | ------ | ------ |------ |------ |
| 79 | 9/26/2018 18:26:32 | 1 | 1.question1? | 3 | …… |
| 79 | 9/26/2018 18:26:32 | 2 |2.question2？| 2 | ……|
| 79 | 9/26/2018 18:26:32 | 3 |3.question3？| 4 | ……|

**after merging**

| number | submitTime | 1.question1?  | 2.question2？ | 3.question3？ | other |
| -------- | -------- | ------ | ------ |------ |------ |
| 1 | 9/26/2018 18:26:32 | 3 | 2| 4 | …… |

Among them, the third column, the fourth column, and the fifth column respectively correspond to the answers to the question.

- The questions answered by each user may not be all the questions, because the questionnaire may jump to different options according to the user's answer, for example, a total of 10 questions questionnaire, the first user answers 1, 3, 5, 7, 8 , 10 questions, but the second user may have answered 1, 2, 4, 5, 9, 10 questions, and it is also possible that the user answers all questions.

- There may be multiple choices in the questionnaire。

## Usage
ExcelPoi project provides [ExcelPoi.jar](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/ExcelPoi.jar)

### Run environment
- JRE1.8 or JDK1.8 or above

### Run command
```java -jar ExcelPoi.jar fileUrl saveUrl saveFileName```
- fileUrl represents the original data path, such as /home/root/question.xlsx or C:\question.xlsx
- saveUrl represents the path saved after data processing, such as /home/root/ or C:\
- saveFileName represents file name that is finally saved. Optional.

## API reference
- [How to create a new workbook](http://poi.apache.org/components/spreadsheet/quick-guide.html#NewWorkbook)
- [How to create a sheet](http://poi.apache.org/components/spreadsheet/quick-guide.html#NewSheet)
- [How to create cells](http://poi.apache.org/components/spreadsheet/quick-guide.html#CreateCells)
- [Iterate over rows and cells](http://poi.apache.org/components/spreadsheet/quick-guide.html#Iterator)
- [Getting the cell contents](http://poi.apache.org/components/spreadsheet/quick-guide.html#CellContents)
- [Reading and writing](http://poi.apache.org/components/spreadsheet/quick-guide.html#ReadWriteWorkbook)

## Project creation
### Build with maven
- Add poi dependencies to support.xls format
```
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->
<dependency>
   <groupId>org.apache.poi</groupId>
   <artifactId>poi</artifactId>
   <version>3.17</version>
</dependency>
```

- Add poi-ooxml dependency to support .xlsx format
```
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>
   <groupId>org.apache.poi</groupId>
   <artifactId>poi-ooxml</artifactId>
   <version>3.17</version>
</dependency>
```

- If the exception of  ```java.lang.NoClassDefFoundError: org/apache/xmlbeans/XmlObject``` occurs, you still need to add the xmlbeans dependency.
```
<!-- https://mvnrepository.com/artifact/org.apache.xmlbeans/xmlbeans-->
<dependency>
   <groupId>org.apache.xmlbeans</groupId>
   <artifactId>xmlbeans</artifactId>
   <version>3.0.2</version>
</dependency>
```

### Support.xls and .xlsx formats
If origin file is an .xls format, use the HSSFWorkbook to create a Workbook, and if it is an .xlsx format, create it with XSSFWorkbook.

### POI concept explanation
- Workbook：Excel workbook  
- Sheet：Excel sheet
- Row：Excel row
- Cell：Excel cell

As shown below:

![Workbook](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/workbook_eng.png)
