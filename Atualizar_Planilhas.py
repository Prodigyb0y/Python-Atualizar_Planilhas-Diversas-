import time
import win32com.client as win32

def atualizar_planilha(caminho_planilha, tempo_espera=10):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    print(f"\nIniciando Atulização Planilha {caminho_planilha}...")
    try:
        workbook = excel.Workbooks.Open(caminho_planilha)
        time.sleep(tempo_espera)
        print("Caminho aberto...")  

        # Atualiza as querys em todas as planilhas
        workbook.RefreshAll()
        print("Atualizando querys...")
        excel.CalculateUntilAsyncQueriesDone()
        print("Consultas atualizadas.")
        time.sleep(2)
      
        try:
            workbook.Save()
            
            time.sleep(2)
            print("Planilha salva!")
        except Exception as e:
            print(f"Erro ao salvar a planilha: {e}")

    except Exception as e:
        print(f"Erro ao abrir ou processar a planilha: {e}")

    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except Exception as e:
                print(f"Erro ao fechar a planilha: {e}")
        if excel:
            excel.Quit()

    print(f"\nProcesso: {caminho_planilha} finalizado.")


# Lista de planilhas a serem atualizadas (ordem alterada)
planilhas = [
    r"caminho_planilhas",
    r"caminho_planilhas",
]

# Atualiza as planilhas em sequência
for planilha in planilhas:
    atualizar_planilha(planilha)

print("\nProcesso Finalizado com sucesso!")