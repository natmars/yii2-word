<?php
namespace natmars\example;

use natmars\word\Word;
use Yii;
use yii\base\Component;

class Test extends Component
{
    public static function generate()
    {
        $fileName = 'template.docx';
        $templateFullPath = Yii::getAlias('@natmars/example') . '/templates/' . $fileName;
        $outputFullPath = Yii::$app->getRuntimePath() . '/' . $fileName;

        $phpWord = new Word();

        // create a file using a template
        $phpWord->saveFromMultiLineTemplate($templateFullPath, $outputFullPath, [
            'variable_1' => 'value_1',
            'variable_2' => 'value_2',
            'variable_3' => 'value_3',
            'table_1' => [
                ['item_1' => '111', 'item_2' => '112'],
                ['item_1' => '121', 'item_2' => '122'],
                ['item_1' => '131', 'item_2' => '132'],
            ],
            'table_2' => [
                ['item_1' => '211', 'item_2' => '212', 'item_3' => '213'],
                ['item_1' => '221', 'item_2' => '222', 'item_3' => '223'],
                ['item_1' => '231', 'item_2' => '232', 'item_3' => '233'],
            ],
        ]);

        // download file
        $phpWord->downloadTemplate($outputFullPath);
    }
}