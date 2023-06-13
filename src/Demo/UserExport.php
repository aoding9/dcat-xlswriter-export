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
                $row->phone,
                $row->created_at->toDateTimeString(),
            ];
            // dd($rowData);
            return $rowData;
        } catch (Exception $e) {
            throw new Exception('id为' . $row->id . '的记录导出失败，原因：' . $e->getMessage());
        }
    }
}
