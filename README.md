### 简介

dcat扩展：xlswriter导出

示例：


### 安装
`composer require aoding9/laravel-baidu-aip`

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

阿里云镜像（可能安装失败）

`composer config -g repo.packagist composer https://mirrors.aliyun.com/composer/`


### 配置

暂无配置




### 使用
以用户导出为例，首先创建一个UserExport导出类，继承`Aoding9\Dcat\Xlswriter\Export\BaseExport`基类，一般放在app\Admin\Exports目录下
```php
<?php

namespace App\Admin\Exports;

use Aoding9\Dcat\Xlswriter\Export\BaseExport;
use Illuminate\Database\Eloquent\Model;
use Aoding9\Dcat\Xlswriter\Export\Demo\User; // 要导出的模型，用于代码提示

class UserExport extends BaseExport {
    // 定义表头
    public $header = [
        ['column' => 'a', 'width' => 8, 'name' => '序号'],
        ['column' => 'b', 'width' => 8, 'name' => 'id'],
        ['column' => 'c', 'width' => 20, 'name' => '姓名'],
        ['column' => 'd', 'width' => 20, 'name' => '手机号'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    
    ];
    public $fileName = '用户导出表'; // 导出的文件名
    public $tableTitle = '用户导出表'; // 第一行标题
    public $rowHeight = 30; // 行高
    public $headerRowHeight = 50; // 表头行高
    
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
                $row->phone,
                $row->created_at->toDateTimeString(),
            ];
            // dd($rowData);
            return $rowData;
        } catch (\Exception $e) {
            dd('id为' . $row->id . '的记录导出失败，原因：' . $e->getMessage());
        }
    }
}

```

在dcat的`UserController`的grid方法中，添加如下代码：
```php
namespace App\Admin\Controllers;
use Dcat\Admin\Grid;
use Aoding9\Dcat\Xlswriter\Export\Demo\User;
use Aoding9\Dcat\Xlswriter\Export\Demo\UserExport;
class UserController {
    protected function grid()
    {
        return Grid::make(new User(), function (Grid $grid) {
            //添加这行即可
             $grid->export(new UserExport());
        });
    }
}

```
如果map中需要调用关联关系，可以在grid中使用with来预加载关联，从而优化查询。

