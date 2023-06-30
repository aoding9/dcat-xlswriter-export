## 简介

dcat扩展：xlswriter导出

之前用了laravel-excel做数据导出，太耗内存速度也慢，数据量大的时候内存占用容易达到php上限，或者响应超时，换成xlswriter这个扩展来做。

由于xlswriter直接导出的表格不够美观，在实际使用中，往往需要合并单元格和自定义表格样式等，我进行了一些封装，使用更加方便简洁，定义表头和数据的方式也更加直观。

本扩展是(laravel-xlswriter-export)[https://github.com/aoding9/laravel-xlswriter-export]的dcat版本，使用文档直接看laravel-xlswriter-export的readme即可，此处只介绍不同点。

## 安装

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

## 使用

### 1.定义导出类
和laravel-xlswriter-export相同，不同点在于，不传参数将使用dcat中grid的查询构造器作为数据源，另外由于make方法与dcat不兼容，只能使用new的方式。

### 2、在控制器中使用

在grid方法中使用`$grid->export(new UserExport());`即可
```php
//use xxx

protected function grid() {
    return Grid::make(new User(), function(Grid $grid) {
         $grid->export(new UserExport());
        //$grid->export(new UserMergeExport());
    });
}
```