@echo off

rem Ativa o ambiente base do Miniforge3
call conda activate base

rem Executa o script Python em uma janela minimizada
start /min "" python "HostFlow.py"

rem Fecha o terminal
exit
