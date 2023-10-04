# Site-Analysis-Exercise

Um código Python utilizando as bibliotecas Pandas, openpyxl e XlsxWriter. O código analisa as duas tabelas e pega as informações solicitadas de "Results" dos sites de "SiteList". A partir dessas informações, o código retorna um arquivo .xlsx com as informações solicitadas no enunciado do exercício.
Além disso, o código verifica na planilha "Results" diferente casos e mostra no terminal do Python as informações. Vale ressaltar que não será mostrado "Sites que não estão presentes no Results" pois a base de dados é o próprio "Results", como solicitado no exercício.

Um diferencial do código apresentado é que alem do terminal do Python, as informações solicitadas no segundo item serão salvas em diferentes abas do mesmo arquivo Excel, graças a biblioteca XlsxWriter.

Vale ressaltar os seguintes pontos:
1. As duas tabelas de input e o código Python devem estar na mesma pasta.
2. Será necessário baixar as bibliotecas utilizadas na maquina. Para isso utilize os comandos a seguir (utilizando o pip atualizado):
   ```ruby
   pip install pandas
   pip install openpyxl
   pip install XlsxWriter
   ```
