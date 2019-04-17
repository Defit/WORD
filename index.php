<?php include_once './WordWorker.php';?>
<?php include_once './MySQLWord.php';?>
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
            
            $ww->setSectionDefaultStyle(['orientation' => 'portrait'], ['colsNum' => 3]);
            
            $ww->addSectionDefault('three');
            $ww->setTextToSection('three', 'Третья секция');
            $ww->setTextToSection('three', 'Третья секция 2');
            
            $ww->setListStyle(3,'Arial',20);
            $ww->addListItem('first', 'Новый элемент');
            $ww->addListItem('first', 'Новый элемент подсписка', 1);
            
            $ww->setTableStyle(['borderSize' => 1], ['borderColor' => '999999'],
                ['afterSpacing' => 0], ['Spacing' => 0], ['cellMargin' => 0]);
            

            $ww->addTableToSection('first', 'firstT');
            
            $cell_style = [['borderTopSize'=>1], ['borderTopColor' =>'black'],
                 ['borderLeftSize'=>1],['borderLeftColor' =>'black'],['borderRightSize'=>1],
                 ['borderRightColor'=>'black'],['borderBottomSize' =>1],['borderBottomColor'=>'black']];
            $ww->setTextStyle(['spaceAfter' => 0]);
            
            $ww->setTableRowStyle(['exactHeight' => -5]);
            $ww->setTableCellStyle(...$cell_style);
            $ww->selectTable('firstT');
            
//            $ww->addTableRow(-0.5);
//            $ww->addTableCell('Cell 1.1', 2500);
//            $ww->addTableCell('Cell 1.2', 2500);
//            $ww->addTableRow();
//            $ww->addTableCell('Cell 2.1', 2500);
//            $ww->addTableCell('Cell 2.2', 2500);
            
            $msw = new MySQLWord($ww);
            
            $msw->setTextStyle(['size' => '12']);
            $msw->setTextStyle(['Color' => '#008000']);
            $msw->setTableCellStyle(['valign' => 'center']);
            
            $connection_string = array(['host' => 'localhost'], 
                ['login' => 'ugpir'], ['password' => 'J8Wti7'], 
                ['db_name' => 'gpir']);
            
            $sql = "SELECT * FROM plan_stage WHERE ID > ? AND ID < ?";
            
            $msw->generateSelectTable($connection_string, $sql, "ii", 2, 5);
            
            $msw->saveDoc('TEST_DOC');
            
            // http://docs.mirocow.com/doku.php?id=php:docx_doc
            // PARAMS
        ?>
    </body>
</html>































































































































