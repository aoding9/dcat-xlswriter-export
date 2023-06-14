<?php

namespace Aoding9\Dcat\Xlswriter\Export\Demo;

use Aoding9\Dcat\Xlswriter\Export\BaseExport;

// 要导出的模型，用于代码提示

class AreaExport extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 8, 'name' => 'id'],
        ['column' => 'b', 'width' => 10, 'name' => '名称'],
        ['column' => 'c', 'width' => 10, 'name' => '简称'],
        ['column' => 'd', 'width' => 20, 'name' => '地区码'],
    
    ];
    
    public $fileName = '地区导出表'; // 导出的文件名
    public $tableTitle = '地区导出表'; // 第一行标题
    
    public $useTitle=false; // 不使用首行标题
    public $useFreezePanes=false; // 不使用冻结
    public $debug=true; // 开启调试，查看耗时和内存占用情况
    // public $max=500000;
    // public $chunkSize=50000;
    public $chunkSize=2000;
    public $max=10000;
    
    // 将模型字段映射为数组
    public function eachRow($row) {
        return [
            $row->id,
            $row->name,
            $row->short_name,
            (string)$row->area_code, // 不转string会导出科学计数法的格式
        ];
    }
}
