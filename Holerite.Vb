'---------------------------------------------
'                                            '
'   Nome: Miguel Fernandes Santos.           '
'   Data: 29/10/2021                         '
'   Projeto: Holerite. (Didatico)            '
'                                            '
'---------------------------------------------

' VARIAVEIS PARA OS VALORES DAS CAIXAS DE TEXTO.
Dim quantas_hr_extra     As Double
Dim quantas_faltas       As Double
Dim quantos_atrasos      As Double

' VARIAVEIS PARA CALCULO -- POSITIVOS.
Dim salario_dia          As Double
Dim salario_hora         As Double
Dim salario_minuto       As Double
Dim hora_extra           As Double
Dim hora_extra_total     As Double
Dim dsr_hr_extra         As Double

' VARIAVEIS PARA CALCULO -- NEGATIVOS.
Dim falta                As Double
Dim dsr_falta            As Double
Dim atraso               As Double

' VARIAVEIS PARA IMPOSTOS.
Dim base_inss            As Double
Dim INSS                 As Double
Dim inss_a_deduzir       As Double
Dim faixa_inss           As Double
Dim porcentagem_inss     As Double
Dim base_fgts            As Double
Dim fgts                 As Double
Dim base_irrf            As Double
Dim Valor_dependentes    As Double
Dim irrf                 As Double
Dim irrf_deduzir         As Double
Dim porcentagem_irrf     As Double

'VARIAVEIS PARA PEGAR OS VALORES CADASTRADOS.
Dim salario              As Double
Dim hr_extra             As Double
Dim insalubridade        As Double
Dim periculosidade       As Double
Dim dependentes          As Double
Dim v_t                  As Double

' VARIAVEIS PARA VALORES FINAIS
Dim hora_extra_final     As Double
Dim insalubridade_final  As Double
Dim periculosidade_final As Double
Dim V_transporte_final   As Double
Dim INSS_FINAL           As Double
Dim FGTS_FINAL           As Double
Dim IRRF_FINAL           As Double

Dim VENCIMENTO_TOTAL     As Double
Dim DESCONTO_TOTAL       As Double
Dim SALARIO_LIQUIDO      As Double

'FIM DAS VARIAVEIS.

' QUAL FUNCIONARIO.
carregar = "Miguel Fernandes Santos"

' INFORMAÇÕES DO CADASTRO DO FUNCIONARIO.
salario = 5000
hr_extra = 75
insalubridade = 0
periculosidade = 0
dependentes = 1
v_t = 6

' FIM DOS VALORES.

quantas_hr_extra = 5
quantas_faltas = 2
quantos_atrasos = 2

' ------- SALARIO DIA, HORA E MINUTO ------- '

' --- SALARIO DIA --- '
 salario_dia = salario / 30

' --- SALARIO HORA --- '
 salario_hora = salario / 220

' --- SALARIO MINUTO ---'
 salario_minuto = salario_hora / 60


' ------- HORA EXTRA E DSR HORA EXTRA ------- '

' --- HORA EXTRA --- '
 hora_extra = (salario_hora / 100) * hr_extra
 hora_extra_total = salario_hora + hora_extra
 hora_extra_final = hora_extra_total * quantas_hr_extra

' --- DSR HORA EXTRA --- '
 dsr_hr_extra = (hora_extra_final / 26) * 4
 
 
' -------------- FALTA E ATRASO ------------- '
 
' --- FALTA --- '
 falta = salario_dia * quantas_faltas

' --- DSR FALTA --- '
 dsr_falta = salario_dia * quantas_faltas

' --- ATRASO --- '
 atraso = salario_hora * quantos_atrasos
 
 
' ------ INSALUBRIDADE E PERICULOSIDADE ----- '

' --- INSALUBRIDADE --- '
insalubridade_final = (1100 / 100) * insalubridade

' --- PERICULOSIDADE --- '
 periculosidade_final = (salario / 100) * periculosidade


' ------------- VALE TRANSPORTE ------------- '
V_transporte_final = (salario / 100) * v_t


' ------- IMPOSTOS: INSS, FGTS E IRRF ------- '

' --- INSS --- '
base_inss = (salario + insalubridade_final + periculosidade_final + hora_extra_final + dsr_hr_extra) - (falta - dsr_falta - atraso)

If base_inss < 1100 Then

    INSS = (base_inss / 100) * 7.5
    inss_a_deduzir = 0
    porcentagem_inss = 7.5

End If

If base_inss > 1100 And base_inss < 2203.48 Then

    INSS = (base_inss / 100) * 9
    inss_a_deduzir = 16.5
    porcentagem_inss = 9

End If

If base_inss > 2203.48 And base_inss < 3305.22 Then

    INSS = (base_inss / 100) * 12
    inss_a_deduzir = 82.6
    porcentagem_inss = 12

End If

If base_inss > 3305.22 And base_inss < 6433.56 Then

    INSS = (base_inss / 100) * 14
    inss_a_deduzir = 148.71
    porcentagem_inss = 14

End If

If base_inss > 6433.56 Then

    base_inss = 6433.56
    INSS = 751.97
    inss_a_deduzir = 0
    porcentagem_inss = 14

End If

INSS_FINAL = INSS - inss_a_deduzir
' --- FIM INSS --- '

' --- FGTS --- '
base_fgts = salario + insalubridade_final + periculosidade_final + hora_extra_final + dsr_hr_extra - falta - dsr_falta - atraso

fgts = (base_fgts / 100) * 8
FGTS_FINAL = fgts

' --- FIM FGTS --- '

' --- IRRF --- '
Valor_dependentes = 189.59 * dependentes
base_irrf = base_fgts - INSS_FINAL - Valor_dependentes

If base_irrf < 1903.99 Then

    irrf = 0
    irrf_deduzir = 0
    porcentagem_irrf = 0

End If

If base_irrf > 1903.98 And base_irrf < 2826.65 Then

    irrf = (base_irrf / 100) * 7.5
    irrf_deduzir = 142.8
    porcentagem_irrf = 7.5

End If

If base_irrf > 2826.66 And base_irrf < 3751.05 Then

    irrf = (base_irrf / 100) * 15
    irrf_deduzir = 354.8
    porcentagem_irrf = 15

End If

If base_irrf > 3751.06 And base_irrf < 4664.68 Then

    irrf = (base_irrf / 100) * 22.5
    irrf_deduzir = 636.13
    porcentagem_irrf = 22.5

End If

If base_irrf > 4664.69 Then

    irrf = (base_irrf / 100) * 27.5
    irrf_deduzir = 869.36
    porcentagem_irrf = 27.5
    
End If

IRRF_FINAL = irrf - irrf_deduzir


' ---------- SALARIO LIQUIDO ---------- '
SALARIO_LIQUIDO = salario + hora_extra_final + dsr_hr_extra + insalubridade_final + periculosidade_final - falta - dsr_falta - atraso - V_transporte_final - INSS_FINAL - IRRF_FINAL

'FIM DO CODIGO :)

End Sub
