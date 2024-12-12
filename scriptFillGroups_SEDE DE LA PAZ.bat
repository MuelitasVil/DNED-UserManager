@echo off
REM Script para asignar usuarios a los grupos en la sede SEDE DE LA PAZ

REM Llenar grupo: SEDE_estudiante_de@unal.edu.co
gam csv "estudiante_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

REM Llenar grupo: FACULTAD_estepp_de@unal.edu.co
gam csv "estepp_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

REM Llenar grupo: FACULTAD_estepppre_de@unal.edu.co
gam csv "estepppre_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

REM Llenar grupo: PLAN_L001_de@unal.edu.co
gam csv "L001_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

REM Llenar grupo: PLAN_L002_de@unal.edu.co
gam csv "L002_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

REM Llenar grupo: PLAN_L003_de@unal.edu.co
gam csv "L003_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

REM Llenar grupo: PLAN_L004_de@unal.edu.co
gam csv "L004_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

REM Llenar grupo: PLAN_L005_de@unal.edu.co
gam csv "L005_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

REM Llenar grupo: PLAN_L006_de@unal.edu.co
gam csv "L006_de@unal.edu.co.csv" gam update group "~Group Email" add "~Member Role" "~Member Email"

