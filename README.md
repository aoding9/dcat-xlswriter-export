### 简介

dcat扩展：xlswriter导出

之前用了laravel-excel做dcat的数据导出，那玩意太耗内存了，数据量大的时候直接超时卡死，换xlswriter这个扩展来搞。



效果：
![Laravel](https://cdn.learnku.com/uploads/images/202306/08/78338/1EjVb0begV.png!large)

![Laravel](https://cdn.learnku.com/uploads/images/202306/08/78338/PKyLtlX9DV.png!large)

### 安装

首先按文档把xlswriter扩展安装上

https://xlswriter-docs.viest.me/

在phpinfo中确认安装成功后，进行下一步

`composer require aoding9/dcat-xlswriter-export`

因为官方源下载太慢了，国内镜像又有各种问题可能导致安装失败，可以把以下代码添加到composer.json，直接从github安装
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

官方源（速度慢）

`composer config -g repo.packagist composer https://packagist.org`


### 配置

暂无配置




### 使用
以用户导出为例，首先创建一个UserExport导出类，继承`Aoding9\Dcat\Xlswriter\Export\BaseExport`基类，一般放在app\Admin\Exports目录下
```php
<?php
/**
 * @Desc 用户导出
 * @User yangyang
 * @Date 2022/8/9 17:01
 */

namespace Aoding9\Dcat\Xlswriter\Export\Demo;

use Aoding9\Dcat\Xlswriter\Export\BaseExport;
use Exception;
use Illuminate\Database\Eloquent\Model;
use Aoding9\Dcat\Xlswriter\Export\Demo\User; // 要导出的模型，用于代码提示

class UserExport extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 8, 'name' => '序号'],
        ['column' => 'b', 'width' => 8, 'name' => 'id'],
        ['column' => 'c', 'width' => 20, 'name' => '姓名'],
        ['column' => 'd', 'width' => 20, 'name' => '手机号'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    
    ];
    
    public $fileName = '用户导出表'; // 导出的文件名
    public $tableTitle = '用户导出表'; // 第一行标题
    //public $rowHeight = 30; // 行高 可选配置项
    //public $headerRowHeight = 50; // 表头行高 可选配置项
    
    // 将模型字段映射为数组
    public function map($row) {
        // 跳过空行
        if (!$row instanceof Model) {
            return $row;
        }
        // dd($row); // $row就是dcat表格每一行对应的模型实例
        try {
            /** @var User $row 用于代码提示 */
            $rowData = [
                $this->index,      // 自增序号，用于无id时，排查导出失败的行
                $row->id,
                $row->name,
                $row->phone??'',
                $row->created_at->toDateTimeString(),
            ];
            // dd($rowData);
            return $rowData;
        } catch (Exception $e) {
            throw new Exception('id为' . $row->id . '的记录导出失败，原因：' . $e->getMessage());
        }
    }
}


```
如果map中需要调用关联关系，可以在grid中使用with来预加载关联，从而优化查询。
仓库中包含UserExport的demo,如果你的已有users表和User模型，可以尝试使用demo进行导出测试

