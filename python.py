from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd, PySimpleGUI as sg
			
url = 'https://portalacademico.ubec.edu.br/framehtml/web/app/edu/PortalEducacional/'
url_login = f'{url}login'
url_notas = f'{url}#/notas'
cabecalho = ["Disciplina", "Situação", "Média Final", "1ª Avaliação", "2ª Avaliação", "3ª Avaliação", "Prova Final"]
wait = 60
df2 = pd.DataFrame(columns=cabecalho)

campo_login = '//*[@id="User"]'
campo_senha = '//*[@id="Pass"]'
botao_entrar = '/html/body/div[2]/div[3]/form/div[4]/input'
menu_sanduiche = '//*[@id="show-menu"]'
lista_central_aluno = '//*[@id="EDU_PORTAL_ACADEMICO_CENTRALALUNO"]'
item_notas = '//*[@id="EDU_PORTAL_ACADEMICO_CENTRALALUNO_NOTAS"]'
alterar_curso = '//*[@id="menu-header-items"]/ul/li[3]/span'
ads_23_2 = '//*[@id="divListaCursos"]/div[2]/div[1]'
nome_facul = '//*[@id="MYGRID"]/div/div[3]/table/tbody/tr[1]/td[1]/span'
xpath_tabela = '//*[@id="MYGRID"]/div/div[3]/table'
div_pai_periodos = '//*[@id="divListaCursos"]'

def criar_interface():
	layout = [
		[sg.Table(values=df2.values.tolist(),
				  headings=cabecalho,
				  auto_size_columns=False,
				  justification='center',
				  display_row_numbers=False,
				  num_rows=8,
				  col_widths=[30, 15, 9, 9, 9, 9, 9],
				  key='table')],
		[sg.Text("1. Digite sua matrícula")],
		[sg.InputText(key='matricula')],
		[sg.Text("2. Digite sua senha")],
		[sg.InputText(key='senha',
					  password_char='*')],
		[sg.Text("3. Digite o período letivo")],
		[sg.InputText(key='periodo',
				default_text='2024/1')],
		[sg.Button('Salvar e Executar')]
	]
	return sg.Window('Notas do(a) Aluno(a)', layout, size=(900, 400))

def atualizar_tabela(janela, dados):
	janela['table'].update(values=dados)

def capturar_notas(matricula, senha, periodo):
	chrome_options = ChromeOptions()
	chrome_options.add_argument('--start-maximized')
	# Caso queira ver o Selenium em funcionamento, comente a linha de código abaixo.
	chrome_options.add_argument('headless')
	chrome_service = ChromeService(ChromeDriverManager().install())
	nav = webdriver.Chrome(service=chrome_service, options=chrome_options)

	nav.get(url_login)

	WebDriverWait(nav, wait).until(EC.visibility_of_element_located((By.XPATH, campo_login))).send_keys(matricula)
	nav.find_element(By.XPATH, campo_senha).send_keys(senha)
	nav.find_element(By.XPATH, botao_entrar).click()

	WebDriverWait(nav, wait).until(EC.visibility_of_element_located((By.XPATH, menu_sanduiche)))
	WebDriverWait(nav, wait).until(EC.invisibility_of_element_located((By.ID, 'loading-screen')))
	
	try:
		checagem_div_periodos = nav.find_element(By.XPATH, div_pai_periodos)
	except NoSuchElementException:
		checagem_div_periodos = False

	if checagem_div_periodos is False:
		WebDriverWait(nav, wait).until(EC.element_to_be_clickable((By.XPATH, alterar_curso))).click()
		WebDriverWait(nav, wait).until(EC.visibility_of_element_located((By.XPATH, div_pai_periodos)))

	div_pai = nav.find_element(By.XPATH, div_pai_periodos)
	divs_filhas = div_pai.find_elements(By.XPATH, ".//div")

	for div_filha in divs_filhas:
		if periodo in div_filha.text:
			div_filha.click()
			break
	
	nav.find_element(By.ID, 'btnConfirmar').click()
	WebDriverWait(nav, wait).until(EC.invisibility_of_element_located((By.ID, 'loading-screen')))
	nav.get(url_notas)
	WebDriverWait(nav, wait).until(EC.visibility_of_element_located((By.XPATH, nome_facul)))
	tabela = nav.find_element(By.XPATH, xpath_tabela)

	dados_tabela = []

	for linha in tabela.find_elements(By.TAG_NAME, "tr"):
		dados_linha = []

		for celula in linha.find_elements(By.TAG_NAME, "td"):
			dados_linha.append(celula.text)

		dados_tabela.append(dados_linha)
	nav.quit()
	df = pd.DataFrame(dados_tabela)
	df2 = df.iloc[:, 2:9].copy()
	df2.columns = cabecalho
	return df2

window = criar_interface()

while True:
	event, values = window.read()

	if event == 'Salvar e Executar':
		matricula = values['matricula']
		senha = values['senha']
		periodo = values['periodo']

		df_notas = capturar_notas(matricula, senha, periodo)

		atualizar_tabela(window, df_notas.values.tolist())

	if event == sg.WINDOW_CLOSED:
		break

window.close()