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
            $ww->setProperties();
            $ww->addSection('first');
            
            $params = ['first', 'Это простой текст'];
            
            $ww->setTextToSection(...$params);
            $ww->setListStyle();
            $ww->addListItem('first', 'Новый элемент');
            $ww->addListItem('first', 'Новый элемент подсписка', 1);
            $ww->saveDoc('TEST_DOC', true);
        ?>
    </body>
</html>





















