dim n1,n2,n3
dim resp, audio, vozAtivado
dim numeroMaior

resp=msgbox("Deseja ativar o recurso de voz?", vbquestion+vbyesno,"ATENCAO")
if resp=vbyes then
    vozAtivado = 1
    call carregar_voz
    sub carregar_voz()
    set audio=createobject("SAPI.SPVOICE")
    audio.rate=1
    audio.volume=100
    call calculoNumero
    end sub
else
    vozAtivado = 0
    call calculoNumero
end if

sub calculoNumero()
    n1= cint (inputbox("Digite o primeiro numero","AVISO"))
    n2= cint (inputbox("Digite o segundo numero","AVISO"))
    n3= cint (inputbox("Digite o terceiro numero","AVISO"))

    if n1 > n2 and n1 > n3 then
        numeroMaior = n1
    elseif n2 > n3 and n2 > n1 then
        numeroMaior = n2
    else 
        numeroMaior = n3
    end if
    
    if vozAtivado = 1 then
        audio.speak("O numero maior e " & numeroMaior & "")
    else
        msgbox("O numero maior e: " & numeroMaior & ""),vbquestion+vbokonly,"Numero Maior"
    end if
    call novoTeste
end sub

sub novoTeste()
resp=msgbox("Deseja realizar novo teste?", vbquestion+vbyesno,"ATENCAO")
if resp=vbyes then
    call calculoNumero
    else
    wscript.quit
end if
end sub