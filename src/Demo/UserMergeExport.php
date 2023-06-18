<?php

namespace Aoding9\Dcat\Xlswriter\Export\Demo;

use Aoding9\Dcat\Xlswriter\Export\BaseExport;

// 要导出的模型，用于代码提示

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
     * @Desc 在分块数据插入每行后回调（到下一个分块，则上一分块被销毁）
     * @param $row
     */
    public function afterInsertEachRowInEachChunk($row) {
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
