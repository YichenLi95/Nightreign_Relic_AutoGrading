# <黑环遗物自动评分器>
一键给存档中全部遗物打分，并支持按阈值批量清理（操作存档前务必备份）

---

## ✨ 功能

### 遗物评分
- 通过导入有save editor导出的遗物数据，对存档中所有的遗物评分。评分基于遗物词条在各流派下的强度。

### 遗物批量删除
- 遗物将根据综合评分百分比得出总分，可选择批量删除总分x以下的遗物。（x为自定义输入值，例如60）

---

## 🛠️ 如何使用

### 遗物评分
1. 找到存档文件 
在文件夹导航栏输入`%APPDATA%\Nightreign`，打开<User-ID>文件夹，找到NR0000.sl2
  （注： 需打开显示隐藏项目， 存档文件尾缀一定是.sl2。）  
  
2. 打开`Elden_Ring_Save_Editor.exe`，将存档文件导入。读取完成后点选角色名字，左上角切换至relics页面，此处展示你存档下所有遗物。（如果是英文的，放大窗口后有language选项，可以切换成中文）导出遗物excel。  
  
3. 以管理员运行`JsonCreation.bat`。（运行完成即可关闭）完成后会创建文件夹`relic_scores_out`  
  
4. 打开`relic_score_multi.exe`, 在excel栏导入由save editor导出的遗物excel；在分数json文件夹栏，选择刚刚自动创建的`relic_scores_out`。  
  
5. 点击加载并显示，即可展示所有遗物的评分。

### 删除遗物
1. 在评分器中删除完遗物后，导出删除后的遗物数据  
  
2. 通过save editor，导入修改后的遗物数据，生成新的.sl2存档文件  
  
3. 替换原先的.sl2文件。  
⚠️⚠️⚠️ 任何对存档的操作都应该提前对存档做备份！！！  
⚠️⚠️⚠️ 任何对存档的操作都应该提前对存档做备份！！！  
⚠️⚠️⚠️ 任何对存档的操作都应该提前对存档做备份！！！

---

## 📥 Download
- 从 Releases 下载最新版本：
    https://github.com/YichenLi95/Nightreign_Relic_AutoGrading/releases
  
- 包中含有save editor。Save editor以Elden-Ring-Nightreign-Save-Editor为原型，修改制作而来，感谢开源作者。
    https://github.com/alfizari/Elden-Ring-Nightreign-Save-Editor

---

## ✅ 前置条件
- Windows 10/11  
  
- 需要管理员权限运行 bat  

---

## ⚠️ 免责声明
- 强烈建议先备份原始数据/文件
  
- 不当修改可能导致数据损坏
  
- 若涉及联网/线上服务，修改后的内容可能触发封禁或风控
  
- 使用风险自负

---

## 🛡️ 叠甲
- 评分标准由作者及朋友填写，小杯理解请大佬谅解，有需要可以自行在"RelicScores.xlsx"中修改分值
  
- 评分基于目前较为广泛流传的流派，会随版本更新
  
- 遗物批量删除功能用于给大佬庞大的遗物库瘦身，请谨慎使用，替换存档前切记备份存档

---

## 👥 Contributors


---


## 📝 License



