# cqf_schedule
cqf_schedule是一个用于生成CQF课表的订阅文件的python脚本。

执行以下脚本：
```
python cqf_schedule.py
```
根据`source`文件夹中的Excel文件，在`ics`文件夹中生成对应的中英文课表订阅文件。

在iphone日历中添加订阅文件后，即可在日历中查看课表。

如果需要同步`gitee`上的仓库，请执行以下脚本：
```
python install_hooks.py
```
会自动在`.git/hooks`中添加`post-commit`钩子，每次`push`时会自动同步到`gitee`上, 便于没有梯子订阅日历日程。