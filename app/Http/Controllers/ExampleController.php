<?php

namespace App\Http\Controllers;

use PhpOffice\PhpWord;
use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\Element\Field;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\SimpleType\TblWidth;
use PhpOffice\PhpWord\Element\Link;

class ExampleController extends Controller
{
    /**
     * Create a new controller instance.
     *
     * @return void
     */
    public function __construct()
    {
        //
    }

    //
    public function test()
    {
        $path = resource_path('./input/bb.docx');
        // 生成world 存放目录
        $filePath = resource_path('./output/bb-auto.docx');
        // 声明模板象并读取模板内容
        try {
            $templateProcessor = new TemplateProcessor($path);


            // 页眉页脚
            $templateProcessor->setValue('WBPXXXXX', 'WBP10028');  // 乙方
            $templateProcessor->setValue('YYYY-A001/A002/A003', '2021-A005/A006/A007');  // 乙方
            $templateProcessor->setValue('thankyou', '测试1');  // 乙方

            // 替换模板内容
            $templateProcessor->setValue('title1', '测试1');  // 乙方
            $templateProcessor->setValue('title2', '测试2');  // 乙方
            $templateProcessor->setValue('title3', '测试3');  // 乙方
            $templateProcessor->setValue('title4', '测试4');  // 乙方
            $templateProcessor->setValue('title5', '测试5');  // 乙方
            $templateProcessor->setValue('title6', '测试6');  // 乙方
            $picParam2 = ['path' => resource_path('./input/ccc.png'), 'width' => 400, 'height' => 400];

            $templateProcessor->setImageValue('image1', $picParam2);
            // $templateProcessor->cloneBlock('block_name', 3, true, true);

            $replacements = array(
                array('customer_name' => 'Batman', 'customer_address' => 'Gotham City'),
                array('customer_name' => 'Superman', 'customer_address' => 'Metropolis'),
            );
            $templateProcessor->cloneBlock('block_name', 0, true, false, $replacements);

            // 克隆模板文档中的表格行
            //$templateProcessor->cloneRow('userid', 2);
            //$templateProcessor->setValue('userid#1', 'wangyao王瑶');

            // 固定列，动态行的表格；
//            $values = [
//                ['userid' => 1, 'username' => '王瑶', 'age' => 18, 'userAddress' => 'Gotham City'],
//                ['userid' => 2, 'username' => '顾淞', 'age' => 19, 'userAddress' => 'Metropolis'],
//                ['userid' => 3, 'username' => '顾淞2', 'age' => 19, 'userAddress' => 'Metropolis'],
//                ['userid' => 4, 'username' => '顾淞3', 'age' => 19, 'userAddress' => 'Metropolis'],
//                ['userid' => 5, 'username' => '顾淞4', 'age' => 19, 'userAddress' => 'Metropolis'],
//            ];
//            $templateProcessor->cloneRowAndSetValues('userid', $values);

            // 固定行，动态列的表格

            // This whole line will be replaced by ${title}
            $title = new TextRun();
            $title->addText('This title has been set ', array('bold' => true, 'italic' => true, 'color' => 'blue'));
            $title->addText('dynamically', array('bold' => true, 'italic' => true, 'color' => 'red', 'underline' => 'single'));
            $templateProcessor->setComplexBlock('title', $title);

            // The following will be replaced ${inline}
            $inline = new TextRun();
            $inline->addText('by a red italic text', array('italic' => true, 'color' => 'red'));
            $templateProcessor->setComplexValue('inline', $inline);

            // This paragraph will be replaced with a ${table}
            $table = new Table(array('borderSize' => 12, 'borderColor' => 'green', 'width' => 6000, 'unit' => TblWidth::TWIP));
            $table->addRow();
            $table->addCell(150)->addText('Cell A1');
            $table->addCell(150)->addText('Cell A2');
            $table->addCell(150)->addText('Cell A3');
            $table->addRow();
            $table->addCell(150)->addText('Cell B1');
            $table->addCell(150)->addText('Cell B2');
            //$table->addCell(150)->addText('Cell B3');
            $templateProcessor->setComplexBlock('table', $table);

            // Let’s insert a date field: ${field}
            $field = new Field('DATE', array('dateformat' => 'dddd d MMMM yyyy H:mm:ss'), array('PreserveFormat'));
            $templateProcessor->setComplexValue('field', $field);

            // And here is a link: ${link}
            $link = new Link('https://github.com/PHPOffice/PHPWord');
            $templateProcessor->setComplexValue('link', $link);

            // 生成新的 world
            $templateProcessor->saveAs($filePath);
        } catch (CopyFileException | CreateTemporaryFileException | Exception $e) {

        }
    }


}
