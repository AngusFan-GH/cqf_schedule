# cqf_schedule
cqf_schedule是一个用于生成CQF课表的订阅文件的python脚本。

## 生成订阅文件
执行以下脚本：
```
python schedule.py
```
即可根据`source`文件夹中的Excel文件，在`ics`文件夹中生成对应的中英文课表订阅文件。

## 订阅日历
订阅链接如下：

中文课：
```
https://raw.githubusercontent.com/AngusFan-GH/cqf_schedule/main/ics/CQF_January_2024_Schedule_chinese.ics
```
英文课：
```
https://raw.githubusercontent.com/AngusFan-GH/cqf_schedule/main/ics/CQF_January_2024_Schedule_english.ics
```
P.S. 上述github链接需要梯子才能访问，如果没有梯子，可以使用`gitee`上的链接：
中文课：
```
https://gitee.com/AngusFan/cqf_schedule/raw/master/ics/CQF_January_2024_Schedule_chinese.ics
```
英文课：
```
https://gitee.com/AngusFan/cqf_schedule/raw/master/ics/CQF_January_2024_Schedule_english.ics
```

在日历中通过`URL`的方式添加订阅文件后，即可在日历中查看课表。

## 同步代码和文件到gitee仓库

如果需要同步`gitee`上的仓库，请执行以下脚本：
```
python install_hooks.py
```
会自动在`.git/hooks`中添加`post-commit`钩子，每次提交时会自动同步到`gitee`上, 便于没有梯子订阅日历日程。