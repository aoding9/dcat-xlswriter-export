## 简介

dcat扩展：xlswriter导出

之前用了laravel-excel做数据导出，很耗内存速度也慢，数据量大的时候内存占用容易达到php上限，或者响应超时，换成xlswriter这个扩展来做。

由于直接导出的表格不太美观，经常需要进行合并单元格和自定义表格样式等操作，我对此进行了一些封装，使用更加方便直观。

本扩展是[laravel-xlswriter-export](https://github.com/aoding9/laravel-xlswriter-export) 的dcat版本，使用文档直接看laravel-xlswriter-export的readme即可，此处只介绍不同点。

## 安装

首先按文档把xlswriter扩展安装上

https://xlswriter-docs.viest.me/

在phpinfo中确认安装成功后，进行下一步

`composer require aoding9/dcat-xlswriter-export`

国内composer镜像如果安装失败，请设置官方源

`composer config -g repo.packagist composer https://packagist.org`

官方源下载慢，国内镜像偶尔出问题可能导致安装失败，也可以把以下代码添加到composer.json，直接从github安装

如果无法访问github,可以将url改为gitee：`https://gitee.com/aoding9/dcat-xlswriter-export`

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

由于构造函数传参为空，数据源为other类型，走`buildDataFromOther`获取grid数据，如有特殊需要，可以重写该方法修改获取到的数据

```php
//use xxx

protected function grid() {
    return Grid::make(new User(), function(Grid $grid) {
         $grid->export(new UserExport());
        //$grid->export(new UserMergeExport());
    });
}
```

### 3、通过swoole使用

由于swoole中不能调用`exit()`方法，需要在控制器中直接return下载响应

为此，需要在导出类中将`$useSwoole`属性设为true,然后在dcat控制器中引入`HandleExportIfUseSwoole`，这个trait将重写index方法，以正确地触发下载

```php
// UserExport中添加
public $useSwoole = true;

// UserController中添加
use Aoding9\Dcat\Xlswriter\Export\HandleExportIfUseSwoole;
```

## 版本更新

- v1.2.2 (2023-9-16)
    - download时调用`$this->useSwoole()`判断是否使用了swoole，如果使用了，将返回下载响应，代替默认的exit()
    - 新增`HandleExportIfUseSwoole`，用于swoole访问dcat时，重写控制器的index以返回下载响应