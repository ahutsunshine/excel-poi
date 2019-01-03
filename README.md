## 使用Apache POI 读写Excel，支持.xls和.xlsx格式
### 业务需求
#### 由于业务上需要处理从数据库中导出的Excel，所以选择开源的Apache POI处理
### 业务描述
#### [questionnaire_simple](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/questionnaire_simple.xlsx)文件简要记录了原数据和处理后的数据的样式和字段，在此解释需要处理的核心字段，其余字段均可忽略。[questionnaire_complete](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/questionnaire_complete.xlsx)文件包含了全部的问卷调查样本。
| field          | meaning            |
| ------------ | ---------------- |
| orderNum |   用户提交问卷的编号 |
| titleKey |   问题序号 |
| titleName|   问题 |
| answer |   答案序号 |

### 业务处理
- 原数据中用户提交问卷的编号对应多条记录，回答每一个问题均对应一条记录，需要将多条记录合并成一条。合并后的表格头由原数据表格基本信息和问题组成。例如：

**原数据**

| orderNum | submitTime | titleKey | titleName | answer | other |
| -------- | ------ | ------ | ------ |------ |------ |
| 79 | 1 | 9/26/2018 18:26:32 | 1.贵行2018年6月末表内资产规模大致范围是? | 3 | …… |
| 79 | 2 | 9/26/2018 18:26:32 |2.贵行2018年6月末资产管理（含保本理财）规模大致范围是？| 2 | ……|
| 79 | 3 | 9/26/2018 18:26:32 |3.资管新规后，资管业务面临的最大调整压力是？| 4 | ……|

**合并后**


| 序号 | 提交时间 | 1.贵行2018年6月末表内资产规模大致范围是?  | 2.贵行2018年6月末资产管理（含保本理财）规模大致范围是？ | 3.资管新规后，资管业务面临的最大调整压力是？ | other |
| -------- | -------- | ------ | ------ |------ |------ |
| 1 | 9/26/2018 18:26:32 | 3 | 2| 4 | …… |

其中，第三列，第四列，第五列分别对应问题的答案需要。

- 每个用户回答的问题可能并不是所有的问题，因为问卷可能会根据用户回答的不同进而跳跃到不同的选项中，例如一共10题问卷调查，用户一回答了1，3，5，7，8，10题，但用户二可能回答了1，2，4，5，9，10题，也有可能用户回答了10题。

- 问卷可能会有多选。

## 使用
项目提供了[ExcelPoi.jar](https://github.com/ahutsunshine/ExcelPoi/blob/master/src/main/resources/ExcelPoi.jar)
### 运行环境
- JRE1.8或JDK1.8或以上
### 运行命令
```java -jar ExcelPoi.jar fileUrl saveUrl ```
- fileUrl表示原数据路径，例如/home/root/question.xlsx或C:\question.xlsx
- saveUrl表示数据处理后保存的路径，例如/home/root/或C:\。

## API参考
- [How to create a new workbook](http://poi.apache.org/components/spreadsheet/quick-guide.html#NewWorkbook)
- [How to create a sheet](http://poi.apache.org/components/spreadsheet/quick-guide.html#NewSheet)
- [How to create cells](http://poi.apache.org/components/spreadsheet/quick-guide.html#CreateCells)
- [Iterate over rows and cells](http://poi.apache.org/components/spreadsheet/quick-guide.html#Iterator)
- [Getting the cell contents](http://poi.apache.org/components/spreadsheet/quick-guide.html#CellContents)
- [Reading and writing](http://poi.apache.org/components/spreadsheet/quick-guide.html#ReadWriteWorkbook)

