<?php
/**
 * @Desc 用户导出
 * @User yangyang
 * @Date 2022/8/9 17:01
 */

namespace Aoding9\Dcat\Xlswriter\Export\Demo;

use Aoding9\Dcat\Xlswriter\Export\BaseExport;
use Illuminate\Database\Eloquent\Model;
use Aoding9\Dcat\Xlswriter\Export\Demo\User;

class UserExport extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 8, 'name' => '序号'],
        ['column' => 'b', 'width' => 8, 'name' => 'id'],
        ['column' => 'c', 'width' => 20, 'name' => '姓名'],
        ['column' => 'd', 'width' => 20, 'name' => '手机号'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    
    ];
    public $fileName = '用户导出表';
    public $tableTitle = '用户导出表';
    public $rowHeight = 30;
    public $headerRowHeight = 50;
    
    public function map($row) {
        if (!$row instanceof Model) {
            return $row;
        }
        // dd($row);
        try {
            /** @var User $row */
            $rowData = [
                $this->index,      //自增序号
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
