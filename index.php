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
            $ww->setSection('first');
            $ww->setTextToSection('first', 'Это простой текст');
            $ww->saveDoc('TEST_DOC', true);
        ?>
    </body>
</html>













