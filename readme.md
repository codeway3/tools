# C3的工具集

## BOM_deleter
eclipse项目转到IDEA时，可能存在utf-8 with BOM的编码，IDEA编译运行时会因为BOM加在文件头的签名报错。  
该脚本自动转码删除文件头的签名，将原文件转为utf-8编码。
  
## uuid_deleter
ireport高版本会自动在元素中加入uuid这个属性，使得jrxml文件无法再使用老版本ireport打开。  
该脚本自动删除uuid属性，并输出到新文件中。
