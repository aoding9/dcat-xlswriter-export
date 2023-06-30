<?php

namespace Aoding9\Dcat\Xlswriter\Export\Demo;


use Aoding9\Dcat\Xlswriter\Export\BaseExport;
use Illuminate\Support\Carbon;
use Vtiful\Kernel\Format;

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
        return [
            $this->index,      // 自增序号，绑定在模型中
            $row->id,
            $row->name = \Faker\Factory::create('zh_CN')->name,
            $this->getGender(),
            $row->created_at,
        ];
    }
    
    public $fileName = '用户导出表';     // 导出的文件名
    public $tableTitle = '用户导出表';   // 第一行标题
    public $useFreezePanes = false; // 是否冻结表头
    public $fontFamily = '宋体';
    public $rowHeight = 30;       // 行高
    public $titleRowHeight = 40;  // 首行大标题行高
    public $headerRowHeight = 50; // 表头行高
    
    /**
     * @Desc 在分块数据插入每行后回调（到下一个分块，则上一分块被销毁）
     * @param $row
     */
    public function afterInsertEachRowInEachChunk($row) {
        // 奇数行进行合并，且不合并到有效数据行之外
        if ($this->index % 2 === 1 && $this->getCurrentLine() < $this->completed + $this->startDataRow) {
            // 定义纵向合并范围，范围形如"B1:B2"
            $range1 = "B" . $this->getCurrentLine() . ":B" . ($this->getCurrentLine() + 1);
            $nextRow = $this->getRowInChunkByIndex($this->index + 1);
            
            $ids = $row->id . '---' . ($nextRow ? $nextRow->id : null);
            // mergeCells（范围, 数据, 样式） ，通过第三个参数可以设置合并单元格的字体颜色等
            $this->excel->mergeCells($range1, $ids, $this->getSpecialStyle());
            
            // 横向合并，形如"C3:D3"
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
        // range是合并范围，$this->end是末尾的列名字母，formatHandle指定合并单元格的样式
        return [
            ['range' => "A1:{$this->end}1", 'value' => $this->getTableTitle(), 'formatHandle' => $this->titleStyle],
            ['range' => "A2:A3", 'value' => '序号', 'formatHandle' => $this->getSpecialStyle()],
            ['range' => "B2:B3", 'value' => 'id', 'formatHandle' => $this->headerStyle],
            ['range' => "C2:E2", 'value' => '基本资料', 'formatHandle' => $this->getSpecialStyle()],
        ];
    }
    
    public $specialStyle;
    
    /**
     * 定义个特别的表格样式
     * @return resource
     */
    public function getSpecialStyle() {
        return $this->specialStyle ?: $this->specialStyle = (new Format($this->fileHandle))
            ->background(Format::COLOR_YELLOW)
            ->fontSize(10)
            ->border(Format::BORDER_THIN)
            ->italic()
            ->font('微软雅黑')
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->wrap()
            ->toResource();
    }
    
    // public $specialStyle2;
    // public $specialStyle3;
    // public function getSpecialStyle2() {}
    // public function getSpecialStyle3() {}
    
    /**
     * @Desc 重写插入单元格数据的处理方法，可单独设置某个单元格的样式
     * @param int           $currentLine  单元格行数
     * @param int           $column       单元格列数
     * @param mixed         $data         插入的数据
     * @param string|null   $format       数据格式化
     * @param resource|null $formatHandle 表格样式
     * @return \Vtiful\Kernel\Excel
     * @Date 2023/6/21 23:37
     */
    public function insertCellHandle($currentLine, $column, $data, $format, $formatHandle) {
        // if($this->getCellName($currentLine,$column)==='A4'){ ... } // 根据单元格名称判断
        
        // 筛选出E列，且日期秒数为偶数的单元格
        if ($this->getColumn($column) === 'E' && $data instanceof Carbon) {
            if ($data->second % 2 === 0) {
                // 设置为上面定义好的样式（黄色背景，斜体，微软雅黑，水平垂直居中等）
                $formatHandle = $this->getSpecialStyle();
            }
            $data = $data->toDateTimeString();
        }
        return $this->excel->insertText($currentLine, $column, $data, $format, $formatHandle);
    }
}
