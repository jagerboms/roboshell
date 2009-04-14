@echo off
dir table.*.sql /b > Schema.scl
dir index.*.sql /b >> Schema.scl
dir fkey.*.sql /b >> Schema.scl
dir proc.*.sql /b >> Schema.scl
rem echo {scl=data\data.scl} >> Schema.scl
