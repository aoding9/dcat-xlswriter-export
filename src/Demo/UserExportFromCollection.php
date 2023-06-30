<?php
namespace Aoding9\Dcat\Xlswriter\Export\Demo;

use Aoding9\Dcat\Xlswriter\Export\BaseExport;

class UserExportFromCollection extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 8, 'name' => '序号'],
        ['column' => 'b', 'width' => 8, 'name' => 'id'],
        ['column' => 'c', 'width' => 10, 'name' => '姓名'],
        ['column' => 'd', 'width' => 10, 'name' => '性别'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    
    ];
    public $fileName = '用户导出表';   // 导出的文件名
    public $tableTitle = '用户导出表'; // 第一行标题
    
    // 将模型字段与表头关联
    public function eachRow($row) {
        return [
            $this->index,
            $row['id'],
            $row['name'],
            random_int(0, 1) ? '男' : '女',
            $row['created_at'],
        ];
    }
    
    public $useGlobalStyle=true;
    // 方法2 可以分块获取数据
    public function buildData(?int $page = null, ?int $perPage = null) {
        return collect([
                           ['id' => 1, 'name' => '小白', 'created_at' => now()->toDateString()],
                           ['id' => 2, 'name' => '小红', 'created_at' => now()->toDateString()],
                       ]);
    }
}
