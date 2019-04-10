<?php include_once './WordWorker.php';?>
<!DOCTYPE html>
<!--
To change this license header, choose License Headers in Project Properties.
To change this template file, choose Tools | Templates
and open the template in the editor.
-->
<html>
    <head>
        <meta charset="UTF-8">
        <title></title>
    </head>
    <body>
        <?php
            $ww = new WordWorker;
            $ww->setDocProperties();

            $ww->setTextStyle(['name' => 'Arial'],
                    ['size' => '18'], ['color' => '000000'],
                    ['italic' => false], ['bold' => false],
                    ['align' => 'left'], ['spaceBefore' => 10]);
            
            $ww->setSectionDefaultStyle(['orientation' => 'landscape'],
                    ['marginTop' => 800], ['marginLeft' => 800],
                    ['marginRight' => 800], ['colsNum' => 1],
                    ['pageNumberingStart' => 1]);
            
            
            $ww->addSection('first');
            
            $params = ['first', 'Это простой текст'];
            $ww->setTextToSection(...$params);
            
            $ww->addSectionDefault('two');
            $ww->setTextToSection('two', 'Вторая секция');
            
            $ww->setSectionDefaultStyle(['portrait' => 'portrait']);
            $ww->addSectionDefault('three');
            $ww->setTextToSection('three', 'Третья секция');
            $ww->setTextToSection('three', 'Третья секция 2');
            
            $ww->setListStyle(3,'Arial',20);
            $ww->addListItem('first', 'Новый элемент');
            $ww->addListItem('first', 'Новый элемент подсписка', 1);
            
            $ww->saveDoc('TEST_DOC', true);
        ?>
    </body>
</html>




















































