Este projeto tem como objetivo criar uma ferramenta no Excel que ajude a organizar e reunir informações essenciais para a declaração de Imposto de Renda.  
A proposta é construir um agregador de dados no qual o usuário possa controlar suas entradas de maneira eficiente e validada, com menus de navegação, validações automáticas e funcionalidades extras, como links rápidos.

A solução é totalmente desenvolvida no Excel, com interface prática e recursos que garantem robustez e usabilidade.

### ✨ Funcionalidades:
- Menus laterais e navegação entre abas.
- Ícones com links diretos para LinkedIn e GitHub, com posição fixa mesmo ao alternar entre abas (via VBA).
- Validações automáticas de campos como CPF, CEP, telefone e celular.
- Interface personalizada com paleta de cores e usabilidade intuitiva.

### 🧩 Formatos personalizados aplicados:
- **CPF**: `000.000.000-00`  
- **CEP**: `00000-000`  
- **Telefone**: `(00) 0000-0000`  
- **Celular**: `(00) 00000-0000`

### ⚙️ Código VBA para posicionar ícones:
Pequena macro utilizada para manter os ícones de redes sociais fixos na tela ao navegar pelas abas.

```vba
Sub MoverIconesParaPosicao()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Call MoverIcone(ws, "icon_linkedin", 50, 305)
    Call MoverIcone(ws, "icon_github", 80, 307)
End Sub

Sub MoverIcone(ws As Worksheet, nomeIcone As String, posX As Double, posY As Double)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = nomeIcone Then
            shp.Left = posX
            shp.Top = posY
            Exit For
        End If
    Next shp
End Sub
