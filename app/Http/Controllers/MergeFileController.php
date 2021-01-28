<?php

namespace App\Http\Controllers;

use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Shared\ZipArchive;
use PhpOffice\PhpWord\Settings;

class MergeFileController extends Controller
{
    private $currentPage = 0;  // 当前分页
    private $page = 0; // 插入页数
    private $args = null; // 文本段样式
    private $tmpFiles = []; // 临时文件

    public function test2()
    {
        $file1 = resource_path('./input/1.docx');
        $file2 = resource_path('./input/2.docx');
        try {
            echo $this->joinFile($file1, $file2, 2);
        } catch (Exception $e) {
        }
    }

    /**
     * 合并文件
     *
     * @param URI
     *    文件1地址
     * @param URI
     *    文件2地址
     * @param Numeric
     *    指定插入的页数
     *
     * @return String
     *    新文件的URI
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    public function joinFile($file1, $file2, $page): string
    {
        $S1 = IOFactory::load($file1)->getSections();
        $S2 = IOFactory::load($file2)->getSections();
        $this->page = $page > 0 ? $page - 1 : $page;

        $phpWord = new PhpWord();

        foreach ($S1 as $S) {

            $section = $phpWord->addSection($S->getStyle());

            $elements = $S->getElements();

            //var_export($elements);exit;
            # 逐级读取／写入节点
            $this->copyElement($elements, $section, $S2);
        }

        $F1 = IOFactory::createWriter($phpWord);
        //$path = $_SERVER['DOCUMENT_ROOT'] . __ROOT__ . '/Public/Write/';
        $path = resource_path('./output/');
        if (!is_dir($path)) mkdir($path);
        $filePath = $path . time() . '.docx';
        $F1->save($filePath);
        # 清除临时文件
        foreach ($this->tmpFiles as $P) {
            unlink($P);
        }
        return $filePath;
    }

    /**
     * 逐级读取／写入节点
     *
     * @param Array
     *    需要读取的节点
     * @param PhpOffice\PhpWord\Element\Section
     *    节点的容器
     * @param Array
     *    文档2的所有节点
     * @return mixed
     */
    private function copyElement($elements, &$container, $S2 = null)
    {
        $inEls = [];
        foreach ($elements as $i => $E) {
            # 检查当前页数
            if ($this->currentPage == $this->page && !is_null($S2)) {
                # 开始插入
                foreach ($S2 as $k => $v) {
                    $ELS = $v->getElements();
                    $this->copyElement($ELS, $container);
                }
                # 清空防止重复插入
                $S2 = null;
            }
            $ns = get_class($E);
            $array = explode('\\', $ns);
            $elName = end($array);
            $fun = 'add' . $elName;

            # 统计页数
            if ($elName == 'PageBreak') {
                $this->currentPage++;
            }

            # 合并文本段
            if ($elName == 'TextRun'
                #&& !is_null($S2)
            ) {
                $tmpEls = $this->getTextElement($E);
                if (!is_null($tmpEls)) {
                    $inEls = array_merge($inEls, $tmpEls);
                }

                $nextElName = '';

                if ($i + 1 < count($elements)) {
                    $nextE = $elements[$i + 1];
                    $nextClass = get_class($nextE);
                    $array = explode('\\', $nextClass);
                    $nextElName = end($array);
                }

                if ($nextElName == 'TextRun') {
                    # 对比当前和下一个节点的样式
                    if (is_object(end($inEls))) {
                        $currentStyle = end($inEls)->getFontStyle();
                    } else {
                        continue;
                    }

                    $nextEls = $this->getTextElement($nextE);
                    if (is_null($nextEls)) {
                        $nextStyle = new Font();
                    } else {
                        $nextStyle = current($nextEls)->getFontStyle();
                    }

                }
            }

            # 设置参数
            $a = $b = $c = $d = $e = null;
            @list($a, $b, $c, $d, $e) = $this->args;
            $newEl = $container->$fun($a, $b, $c, $d, $e);
            $this->setAttr($elName, $newEl, $E);

            #$inEls = [];
            if (method_exists($E, 'getElements')
                && $elName != 'TextRun'
            ) {
                $inEls = $E->getElements();
            }
            if (method_exists($E, 'getRows'))
                $inEls = $E->getRows();
            if (method_exists($E, 'getCells'))
                $inEls = $E->getCells();

            if (count($inEls) > 0) {
                $this->copyElement($inEls, $newEl);
                $inEls = [];
                $this->args = null;
            }

        }

        //return $pageIndex;
        $this->currentPage;
    }

    /**
     * 获取Text节点
     */
    private function getTextElement($E)
    {
        $elements = $E->getElements();
        $result = [];
        foreach ($elements as $inE) {
            $ns = get_class($inE);
            $array = explode('\\', $ns);
            $elName = end($array);
            if ($elName == 'Text') {
                $inE->setPhpWord(null);
                $result[] = $inE;

            } elseif (method_exists($inE, 'getElements')) {
                $inResult = $this->getTextElement($inE);
            }
            if (isset($inResult) && !is_null($inResult))
                $result = array_merge($result, $inResult);
        }
        return count($result) > 0 ? $result : null;
    }

    private function setAttr($elName, &$newEl, $E)
    {
        switch (strtolower($elName)) {
            case 'footnote':
                $newEl->setReferenceId($E->getReferenceId());
                break;
            case 'formfield':
                $newEl->setName($E->getName());
                $newEl->setDefault($E->getDefault());
                $newEl->setValue($E->getValue());
                $newEl->setEntries($E->getEntries());
                break;
            case 'object':
                $newEl->setImageRelationId($E->getImageRelationId());
                $newEl->setObjectId($E->getObjectId());
                break;
            case 'sdt':
                $newEl->setValue($E->getValue());
                $newEl->setListItems($E->getListItems());
                break;
            case 'table':
                $newEl->setWidth($E->getWidth());
                break;
        }

    }
}
