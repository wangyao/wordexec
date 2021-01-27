<?php

namespace App\Http\Controllers;

use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\TemplateProcessor;

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

            // 替换模板内容
            $templateProcessor->setValue('title1', '测试1');  // 乙方
            $templateProcessor->setValue('title2', '测试2');  // 乙方
            $templateProcessor->setValue('title3', '测试3');  // 乙方
            $templateProcessor->setValue('title4', '测试4');  // 乙方
            $templateProcessor->setValue('title5', '测试5');  // 乙方
            $templateProcessor->setValue('title6', '测试6');  // 乙方
            $picParam2 = ['path' => resource_path('./input/ccc.png'), 'width' => 400, 'height' => 400];
            $templateProcessor->setImageValue('image1', $picParam2);
            // 生成新的 world
            $templateProcessor->saveAs($filePath);
        } catch (CopyFileException | CreateTemporaryFileException $e) {

        }
    }
}
