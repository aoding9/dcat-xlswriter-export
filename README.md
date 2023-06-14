### 简介

dcat扩展：xlswriter导出

之前用了laravel-excel做dcat的数据导出，太耗内存速度也慢，数据量大的时候容易超时卡死，换xlswriter这个扩展来搞，分块查询处理，及时销毁已插入的数据，减少内存占用。

**chunk=2000,导出1万条**

![dcat扩展：xlswriter导出，自定义合并单元格，自定义列宽，自定义样式](https://cdn.learnku.com/uploads/images/202306/14/78338/pG9P3d6acx.png!large)

**chunk=50000 导出50万条**

![dcat扩展：xlswriter导出，自定义合并单元格，自定义列宽，自定义样式](https://cdn.learnku.com/uploads/images/202306/14/78338/9WcH75QDKd.png!large)

**导出类简单示例**

![dcat扩展：xlswriter导出，首行合并标题，自定义列宽，表头冻结，composer安装使用](https://cdn.learnku.com/uploads/images/202306/14/78338/PdJ3c08vXO.png!large)


![Laravel](https://cdn.learnku.com/uploads/images/202306/08/78338/1EjVb0begV.png!large)

![Laravel](https://cdn.learnku.com/uploads/images/202306/08/78338/PKyLtlX9DV.png!large)

**导出类自定义合并，自定义样式**


![dcat扩展：xlswriter导出，自定义合并单元格，自定义列宽，自定义样式](https://cdn.learnku.com/uploads/images/202306/15/78338/XqCcdRjK0H.png!large)


![dcat扩展：xlswriter导出，自定义合并单元格，自定义列宽，自定义样式](https://cdn.learnku.com/uploads/images/202306/15/78338/VA8qPNr1kR.png!large)

### 安装

首先按文档把xlswriter扩展安装上

https://xlswriter-docs.viest.me/

在phpinfo中确认安装成功后，进行下一步

`composer require aoding9/dcat-xlswriter-export`

国内composer镜像如果安装失败，请设置官方源

`composer config -g repo.packagist composer https://packagist.org`

因为官方源下载慢，国内镜像又有各种问题可能导致安装失败，也可以把以下代码添加到composer.json，直接从github安装
```json
{
  "repositories": [
    {
      "type": "vcs",
      "url": "https://github.com/aoding9/dcat-xlswriter-export"
    }
  ]
}
```


### 配置

暂无配置


### 使用

以用户导出为例，首先创建一个UserExport导出类，继承`Aoding9\Dcat\Xlswriter\Export\BaseExport`基类，一般放在app\Admin\Exports目录下
```php
<?php
namespace Aoding9\Dcat\Xlswriter\Export\Demo;
use Aoding9\Dcat\Xlswriter\Export\BaseExport;
class UserExport extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 8, 'name' => '序号'],
        ['column' => 'b', 'width' => 8, 'name' => 'id'],
        ['column' => 'c', 'width' => 20, 'name' => '姓名'],
        ['column' => 'd', 'width' => 10, 'name' => '性别'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    
    ];
    
    public $fileName = '用户导出表'; // 导出的文件名
    public $tableTitle = '用户导出表'; // 第一行标题
    
    // 将模型字段与表头关联
    public function eachRow($row) {
            /** @var User $row 用于代码提示 */
            return [
                $this->index,
                $row->id,
                $row->name,
                random_int(0, 1) ? '男' : '女',
                $row->created_at->toDateTimeString(),
            ];
    }
}
```

***合并单元格的示例：***
```php
<?php
namespace Aoding9\Dcat\Xlswriter\Export\Demo;
use Aoding9\Dcat\Xlswriter\Export\BaseExport;
class UserMergeExport extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 10, 'name' => '序号'],
        ['column' => 'b', 'width' => 10, 'name' => 'id'],
        ['column' => 'c', 'width' => 10, 'name' => '姓名'],
        ['column' => 'd', 'width' => 10, 'name' => '性别'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    ];

    public function getGender() {
        return random_int(0, 1) ? '男' : '女';
    }
    // 处理每行的模型，使其对应到表头
    public function eachRow($row) {
        /** @var User $row 用于代码提示 */
        return [
            $this->index,      // 自增序号，绑定在模型中
            $row->id,
            $row->name,
            $this->getGender(),
            $row->created_at->toDateTimeString(),
        ];
    }
    public $fileName = '用户导出表';     // 导出的文件名
    public $tableTitle = '用户导出表';   // 第一行标题
    public $useFreezePanes = false; // 是否冻结表头
    public $fontFamily = '宋体';
    public $rowHeight = 30;       // 行高 可选配置项
    public $titleRowHeight = 40;  // 行高 可选配置项
    public $headerRowHeight = 50; // 表头行高 可选配置项
    /**
     * @Desc 插入一个分块的数据后回调
     * @param $row
     */
    public function afterInsertOneChunkNormalData($row) {
        /** @var User $row */
        // 奇数行进行合并
        // 不合并到数据行之外
        if ($this->index % 2 === 1 && $this->getCurrentLine() < $this->completed + $this->startDataRow) {
            $range1 = "B" . $this->getCurrentLine() . ":B" . ($this->getCurrentLine() + 1);
            $nextRow = $this->getRowByIndex($this->index + 1);

            $ids = $row->id . '---' . ($nextRow ? $nextRow->id : null);
            $this->excel->mergeCells($range1, $ids);

            $range2 = "C" . $this->getCurrentLine() . ":D" . $this->getCurrentLine();
            $nameAndGender = $row->name . "---" . $this->getGender();
            $this->excel->mergeCells($range2, $nameAndGender);
        }
    }

    public function setHeaderData() {
        parent::setHeaderData();
        // 把表头放到第三行，第二行留空用于合并
        $this->headerData->put(2, $this->headerData->get(1));
        $this->headerData->put(1, []);
        return $this;
    }

    /**
     * @Desc 插入数据完成后进行合并
     * @return array[]
     */
    public function mergeCellsAfterInsertData() {
        return [
            ['range' => "A1:{$this->end}1", 'value' => $this->getTableTitle(), 'formatHandle' => $this->titleStyle],
            ['range' => "A2:A3", 'value' => '序号', 'formatHandle' => $this->headerStyle],
            ['range' => "B2:B3", 'value' => 'id', 'formatHandle' => $this->headerStyle],
            ['range' => "C2:E2", 'value' => '基本资料', 'formatHandle' => $this->headerStyle],
        ];
    }
}

```

如果map中需要调用关联关系，可在grid中使用with来预加载关联，从而优化查询。

仓库中包含3个导出类的demo,如果你已有users表和dcat后台控制器，可以尝试使用demo进行导出测试，设置`$debug=true;`即可查看导出的耗时和内存占用。

为了方便自定义排版和修改数据，基类属性和方法都为public，方便子类重写



