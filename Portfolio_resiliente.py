# Importação dos módulos/bibliotecas necessários

import datetime
import traceback
from tkinter import DISABLED, WORD, Button, Tk, Toplevel, Text, END
from tkinter import simpledialog, messagebox
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

import numpy as np
import pandas as pd
import yfinance as yf
import os


class CarteiraApp:

    # Configuração inicial da janela principal
    def __init__(self, root):
        self.root = root
        self.root.title("Análise de Carteira")

        # Variáveis para armazenar informações da carteira
        self.qtd_acoes_cliente = None
        self.acoes_cliente = []
        self.valores_acoes = []
        self.pesos = []
        self.total_investido = None
        self.carteira = None
        self.portfolio_atualizado = None
        self.acao_pretendida = None
        self.retorno_individual = None
        self.valor_pretendido = None
        self.sharpe_nova_acao = None
        self.n_sharpe_carteira = None
        self.risco_anualizado = None
        self.n_acoes_cliente = []
        self.n_valores_acoes = []
        self.n_total_investido = None
        self.n_pesos = None
        self.n_dp_portfolio_anualizado = None
        self.data_inicio = None
        self.data_fim = None

        # Ajuste para centralizar a janela
        largura_janela = 800
        altura_janela = 600

        largura_tela = root.winfo_screenwidth()
        altura_tela = root.winfo_screenheight()

        posicao_x = (largura_tela - largura_janela) // 2
        posicao_y = (altura_tela - altura_janela) // 2

        root.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")

        # Criação dos widgets na janela principal
        self.create_widgets()

    def create_widgets(self):
        # Texto de instruções
        instrucoes = (
            "Fruto de dicas ou de recomendações formais, o investimento em ações pode apresentar oportunidades de "
            "retornos atraentes. Entretanto, dicas de investimentos vistas na internet ou mesmo recomendações de "
            "compra de ações feitas por casas de análises ou analistas certificados têm como objetivo alcançar o maior "
            "número de pessoas possível,fazendo com que investidores adicionem a carteira uma ação que aumenta o risco "
            "total do portfólio sem, necessariamente, aumentar o retorno esperado.\n\n"
            "Assim, o objetivo desta aplicação é auxiliar o investidor a decidir se faz sentido adicionar determinada "
            "ação à carteira atual.\n\n"
            "Ao clicar em Avançar, você será convidado a informar a quantidade de empresas que você tem em carteira. "
            "Em seguida, uma janela se abrirá para que você insira o ticker das ações de uma das empresas da sua "
            "carteira. Ao clicar em Ok, outra janela surgirá para você inserir a posição financeira correspondente. "
            "Novamente surgirá a janela para você inserir o ticker de mais umas das ações da sua carteira e assim por "
            "diante, até que você informe todas as empresas da sua carteira.\n\n"
            "Após inserir os tickers de todas as ações com as respectivas posições financeiras, o sistema irá baixar"
            "as cotações históricas e isso poderá levar alguns minutos. Assim que o download dos dados for concluído, "
            "uma mensagem será exibida na tela. "
            "Ao final, depois de informar a ação que você deseja adquirir, clique no botão 'Gerar Relatório'."
        )

        texto_instrucoes = Text(self.root, height=20, width=100, wrap=WORD, font=("Calibri", 14),
                                bg=self.root.cget("bg"))
        texto_instrucoes.insert(END, instrucoes)
        texto_instrucoes.config(state=DISABLED)  # Impede que o usuário edite o texto
        texto_instrucoes.pack(pady=10)

        # Botão para avançar
        avancar_button = Button(self.root, text="Avançar", command=self.coletar_qtd_acoes)
        avancar_button.pack(pady=10)

        # Botão para gerar relatório
        relatorio_button = Button(self.root, text="Gerar Relatório", command=self.gerar_relatorio)
        relatorio_button.pack(pady=10)

    def coletar_qtd_acoes(self):
        # Verifica se já existe o botão "Avançar" e destroi se existir
        avancar_button = self.root.winfo_children()[1]
        if avancar_button:
            avancar_button.destroy()

        # Solicita ao usuário a quantidade de ações
        qtd_acoes_cliente = simpledialog.askstring("Quantidade de Ações", "Quantas empresas fazem parte da "
                                                                          "sua carteira de ações atual?",
                                                   parent=self.root)
        if qtd_acoes_cliente:
            self.qtd_acoes_cliente = int(qtd_acoes_cliente)

        # Coleta as informações de cada ação
        for _ in range(self.qtd_acoes_cliente):
            ticker = simpledialog.askstring("Ticker da Ação", "Insira o ticker das ações de "
                                                              "uma das empresas da sua carteira (por exemplo: PETR4). "
                                                              "Certifique-se de inserir o ticker correto.",
                                            parent=self.root)
            valor_acao = simpledialog.askfloat("Valor da Ação",
                                               "Informe o valor (R$) da posição financeira atual, sem pontos ou "
                                               "vírgulas. Não é preciso informar os centavos.\n"
                                               "Para uma posição de dez mil e quinhentos reais e cinquenta centavos, "
                                               "insira '10500'.",
                                               parent=self.root)
            self.acoes_cliente.append(ticker)
            self.valores_acoes.append(valor_acao)

            # Calcula os pesos de cada ação informada
            total_investido = sum(self.valores_acoes)
            self.pesos = [valor_acao / total_investido for valor_acao in self.valores_acoes]

        # Criar um novo botão "Avançar"
        avancar_button = Button(self.root, text="Avançar", command=self.continuar_fluxo)
        avancar_button.pack(pady=10)

        # Exibe mensagem informativa
        messagebox.showinfo("Aviso", "O sistema irá baixar os dados das cotações históricas. Assim que o "
                                     "download for concluído, uma nova mensagem será exibida. Clique em OK e aguarde.")

        # Continua o fluxo
        self.continuar_fluxo()

    def continuar_fluxo(self):
        # Adicionar lógica para continuar o fluxo, como baixar dados da yfinance, calcular métricas etc.
        pass

        # Baixar dados da yfinance
        self.baixar_dados_yfinance()

        # Calcular métricas da carteira
        self.calcular_metricas_carteira()

        # Adicionar ação à carteira
        self.adicionar_acao_carteira()

        # Calcular métricas da carteira atualizada
        self.calcular_metricas_carteira_atualizada()

        # Comparar o sharpe ratio
        self.comparar_sharpe_ratios()

    def baixar_dados_yfinance(self):
        try:

            # Calcula as datas de início e fim com base na data atual
            data_atual = datetime.datetime.today().date()
            data_inicio = (data_atual - pd.DateOffset(months=12)).strftime('%Y-%m-%d')
            data_fim = (data_atual - pd.DateOffset(days=1)).strftime('%Y-%m-%d')

            # Baixa dados da yfinance
            carteira_cliente = [acao.upper() + '.SA' for acao in self.acoes_cliente]
            self.carteira = yf.download(carteira_cliente, start=data_inicio, end=data_fim)['Adj Close']

            # Exibe mensagem de sucesso
            messagebox.showinfo("Sucesso", "Dados baixados com sucesso!")

        except Exception as e:
            print(f"Erro ao baixar dados da yfinance: {e}")
            traceback.print_exc()

    # Módulo de cálculo das métricas de risco e retorno da atual carteira do usuário da aplicação
    def calcular_metricas_carteira(self):
        try:
            # Calcula as métricas da carteira
            retornos_individuais = self.carteira.pct_change()
            retornos_individuais.dropna(inplace=True)
            retorno_carteira = (retornos_individuais * self.pesos).sum(axis=1)
            retorno_portfolio = pd.DataFrame()
            retorno_portfolio['Retornos'] = retorno_carteira

            retorno_portfolio_acumulado = pd.DataFrame((1 + retorno_portfolio).cumprod())
            N = len(retorno_portfolio_acumulado)
            retorno_anualizado = retorno_portfolio_acumulado.iloc[-1,] ** (252 / N) - 1
            self.rentabilidade_acumulada = float(round((retorno_portfolio_acumulado.iloc[-1,] - 1), 4))
            self.retorno_anualizado = float(round(retorno_anualizado, 4))

            # Calcula o desvio padrão anualizado dos ativos
            dp_anualizado = retornos_individuais.std() * np.sqrt(252)

            # Calcula a matriz de covariância anualizada
            cov_matrix = retornos_individuais.cov() * 252

            # Calcula o desvio padrão do portfólio
            dp_portfolio_anualizado = np.sqrt(np.dot(np.array(self.pesos).T, np.dot(cov_matrix, self.pesos)))

            sharpe_carteira = (retorno_anualizado - 0.13) / dp_portfolio_anualizado
            self.sharpe_carteira = round(float(sharpe_carteira.iloc[0]), 4)

            self.risco_anualizado = round(dp_portfolio_anualizado, 4)

        except Exception as e:
            # Modificação para imprimir o traceback no console
            print(f"Erro ao calcular métricas da carteira: {e}")
            traceback.print_exc()

    # Coleta informações da ação que o usuário pretende adicionar à carteira e calcular as métricas de risco e retorno
    def adicionar_acao_carteira(self):

        self.data_inicio = (datetime.date.today() - pd.DateOffset(months=12)).strftime('%Y-%m-%d')
        self.data_fim = (datetime.date.today() - pd.DateOffset(days=1)).strftime('%Y-%m-%d')

        # Uso do wait_window para esperar a caixa de diálogo ser fechada
        self.acao_pretendida = simpledialog.askstring("Ação Pretendida",
                                                      "Insira o ticker da ação que representa a empresa que "
                                                      "você deseja adicionar à sua carteira (por exemplo: VALE3)")
        self.valor_pretendido = simpledialog.askfloat("Valor Pretendido",
                                                      "Qual o valor (R$) que você deseja investir nela? Digite "
                                                      "o valor sem pontos ou vírgulas. Não é preciso informar os "
                                                      "centavos.",
                                                      parent=self.root)

        nova_acao = yf.download(self.acao_pretendida.upper() + '.SA', start=self.data_inicio, end=self.data_fim)[
            ('Adj Close')]
        self.retorno_individual = nova_acao.pct_change()
        self.retorno_individual.dropna(inplace=True)

        # Calcula métricas da nova ação
        retorno_n_acao_acumulado = (1 + self.retorno_individual).cumprod()
        N = len(retorno_n_acao_acumulado)
        retorno_n_acao_anualizado = retorno_n_acao_acumulado.iloc[-1,] ** (252 / N) - 1

        # Calcula o desvio padrão anualizado da nova ação
        dp_n_acao_anualizado = self.retorno_individual.std() * np.sqrt(252)

        # Calcula o Sharpe Ratio da nova ação
        self.sharpe_nova_acao = round((retorno_n_acao_anualizado - 0.13) / dp_n_acao_anualizado, 4)

        # Atribui os valores à instância da classe
        self.dp_n_acao_anualizado = round(dp_n_acao_anualizado, 4)
        self.retorno_n_acao_anualizado = round(retorno_n_acao_anualizado, 4)
        self.retorno_n_acao_acumulado = float(round((retorno_n_acao_acumulado.iloc[-1,] - 1), 4))

    # Módulo para calcular as métricas da carteira com a nova ação pretendida pelo usuário
    def calcular_metricas_carteira_atualizada(self):
        try:

            self.n_acoes_cliente = self.acoes_cliente + [self.acao_pretendida]
            self.n_valores_acoes = self.valores_acoes + [self.valor_pretendido]
            self.n_total_investido = sum(self.n_valores_acoes)
            self.n_pesos = [valor_acao / self.n_total_investido for valor_acao in self.n_valores_acoes]

            # Criação de um portfólio para calcular a correlação dos retornos da nova ação com os retornos da carteira
            # já existente
            self.portfolio_proposto = pd.DataFrame()
            retornos_individuais = self.carteira.pct_change()
            retornos_individuais.dropna(inplace=True)
            self.portfolio_proposto['retorno carteira'] = (retornos_individuais * self.pesos).sum(axis=1)
            self.portfolio_proposto['retorno nova ação'] = self.retorno_individual

            # Criação de um portfólio com todos os ativos originais mais a nova ação
            self.portfolio_atualizado = self.carteira.pct_change()
            self.portfolio_atualizado[self.acao_pretendida] = self.retorno_individual
            self.portfolio_atualizado.dropna(inplace=True)

            n_retorno_carteira = (self.portfolio_atualizado * self.n_pesos).sum(axis=1)

            retorno_n_portfolio = pd.DataFrame()
            retorno_n_portfolio['Retornos'] = n_retorno_carteira
            retorno_n_portfolio_acumulado = pd.DataFrame((1 + retorno_n_portfolio).cumprod())

            # Cálculo do retorno anualizado da nova carteira
            N = len(retorno_n_portfolio_acumulado)
            n_retorno_anualizado = retorno_n_portfolio_acumulado.iloc[-1,] ** (252 / N) - 1

            # Matriz de covariância da nova carteira anualizada
            n_cov_matrix = self.portfolio_atualizado.cov() * 252

            # Desvio padrão do portfólio
            self.n_dp_portfolio_anualizado = round(
                np.sqrt(np.dot(np.array(self.n_pesos).T, np.dot(n_cov_matrix, self.n_pesos
                                                                ))), 4)

            n_sharpe_carteira = (n_retorno_anualizado - 0.13) / self.n_dp_portfolio_anualizado
            self.n_sharpe_carteira = round(float(n_sharpe_carteira.iloc[0]), 4)

        except Exception as e:
            print(f"Erro ao calcular métricas da carteira atualizada: {e}")
            traceback.print_exc()

    # Módulo dedicado à comparação da métrica chave do programa
    def comparar_sharpe_ratios(self):
        try:
            # Cálculo da correlação entre os retornos da nova ação com os retornos da carteira já existente
            p = self.portfolio_proposto['retorno carteira'].corr(self.retorno_individual)
            self.correlacao = round(p, 4)

            # Cálculo do segundo elemento da regra geral
            self.segundo_elemento = round((self.sharpe_carteira * p), 4)

            if self.sharpe_nova_acao > self.segundo_elemento:
                self.resultado_teste = (
                    f"Considerando que o retorno ajustado ao risco de {self.acao_pretendida.upper()} é "
                    f"{self.sharpe_nova_acao}, \nportanto maior do que o retorno ajustado ao risco da sua "
                    f"carteira de ações atual, \nmultiplicado pelo coeficiente de correlação entre "
                    f"{self.acao_pretendida.upper()} e a sua carteira, \nque é {self.segundo_elemento}, "
                    "existe benefício em adicioná-la à sua carteira.\n"
                )
            else:
                self.resultado_teste = (
                    f"Considerando que o retorno ajustado ao risco de {self.acao_pretendida.upper()} é "
                    f" {self.sharpe_nova_acao},\nportanto menor do que o retorno ajustado ao risco da sua "
                    f"carteira de ações atual, \nmultiplicado pelo coeficiente de correlação entre "
                    f"{self.acao_pretendida.upper()} e a sua carteira, \nque é {self.segundo_elemento}, "
                    "NÃO existe benefício em adicioná-la à sua carteira.\n"
                )

        except Exception as e:
            print(f"Erro ao comparar Sharpe Ratios: {e}")
            traceback.print_exc()

    # Módulo dedicado à disponibilização de informações ao usuário do programa
    def gerar_relatorio(self):
        # Criar uma nova janela para o relatório
        relatorio_window = Toplevel(self.root)
        relatorio_window.title("Relatório")
        relatorio_window.lift()

        # Adiciona uma Textbox para mostrar as informações
        texto_relatorio = Text(relatorio_window, height=30, width=80, wrap=WORD, font=("Calibri", 12),
                               bg=relatorio_window.cget("bg"))
        texto_relatorio.pack(pady=10)

        # Adiciona informações ao relatório
        relatorio = " - Portfólio Resiliente - Aqui está o seu relatório - \n\n"

        # Adiciona a seção "Janela de Análise" com as datas
        relatorio += ("Os parâmetros de risco e retorno aqui calculados levam em consideração \no desempenho das ações "
                      "entres os dias:\n")
        relatorio += f"Data de Início: {self.data_inicio};\n"
        relatorio += f"Data de Fim: {self.data_fim}.\n\n"

        relatorio += (
            f"Para o cálculo do retorno ajustado risco foi utilizada uma taxa de 13% \ncomo aproximação da taxa "
            f"livre de risco.\n\n")

        # 1) Dados da carteira com a composição atual
        relatorio += "1) Dados da carteira com a composição ATUAL:\n\n"
        relatorio += f"Quantidade de ações: {self.qtd_acoes_cliente};\n"
        relatorio += f"Ações: {self.acoes_cliente};\n"
        relatorio += f"Valores R$: {self.valores_acoes};\n"
        relatorio += f"Total investido R$: {sum(self.valores_acoes)};\n"
        relatorio += f"Rentabilidade acumulada da carteira no período: {'{:.2%}'.format(self.rentabilidade_acumulada)};\n"
        relatorio += f"Rentabilidade anualizada da carteira no período: {'{:.2%}'.format(self.retorno_anualizado)};\n"
        relatorio += f"Risco anualizado (volatilidade da carteira): {'{:.2%}'.format(self.risco_anualizado)};\n"
        relatorio += f"Sharpe (retorno ajustado ao risco) da carteira: {self.sharpe_carteira}.\n\n"

        # 2) Informações sobre a ação pretendida
        relatorio += "2) Informações sobre a ação pretendida:\n\n"
        relatorio += f"Ação: {self.acao_pretendida};\n"
        relatorio += f"Valor Pretendido R$: {self.valor_pretendido};\n"
        relatorio += f"Rentabilidade acumuladada da ação no período: {'{:.2%}'.format(self.retorno_n_acao_acumulado)};\n"
        relatorio += f"Rentabilidade anualizada da ação no período: {'{:.2%}'.format(self.retorno_n_acao_anualizado)};\n"
        relatorio += f"Risco anualizado (volatilidade da ação): {'{:.2%}'.format(self.dp_n_acao_anualizado)};\n"
        relatorio += f"Sharpe (retorno ajustado ao risco) da ação: {self.sharpe_nova_acao}.\n\n"

        # 3) Dados da carteira com a ação pretendida adicionada
        relatorio += "3) Dados da Carteira com a ação pretendida adicionada:\n\n"
        relatorio += f"Quantidade de ações: {len(self.acoes_cliente) + 1};\n"
        relatorio += f"Ações: {self.acoes_cliente + [self.acao_pretendida]};\n"
        relatorio += f"Valores R$: {self.n_valores_acoes};\n"
        relatorio += f"Total que passará a ser investido R$: {sum(self.n_valores_acoes)};\n"
        relatorio += f"Risco Anualizado (Volatilidade da Carteira): {'{:.2%}'.format(self.n_dp_portfolio_anualizado)};\n"
        relatorio += f"Sharpe (retorno ajustado ao risco) da carteira COM a nova ação: {self.n_sharpe_carteira}.\n\n"

        # 4) Dados da carteira com a ação pretendida adicionada
        relatorio += "4) Risco x Retorno:\n\n"
        relatorio += (f"Coeficiente de correlação entre os retornos da ação pretendida \ne os retornos da sua atual "
                      f"carteira: {self.correlacao};\n\n")
        relatorio += (f"Produto do Sharpe da carteira atual vezes o coeficiente de correlação: {self.segundo_elemento};"
                      f"\n\n")
        relatorio += f"{self.resultado_teste}\n\n"

        # Adiciona o relatório à Textbox
        texto_relatorio.insert(END, relatorio)
        texto_relatorio.config(state=DISABLED)  # Impede que o usuário edite o texto

        # Salva o relatório em PDF
        pdf_filename = "Relatorio_Portifólio Resiliente.pdf"
        pdf_path = self.salvar_relatorio_pdf(relatorio, pdf_filename)

                # Exibe mensagem ao usuário
        messagebox.showinfo("Relatório Salvo", f"O relatório está sendo exibido em uma "
                                               f"nova janela e também foi salvo na sua pasta de downloads:\n{pdf_path}")

    def salvar_relatorio_pdf(self, relatorio, pdf_filename):
        # Obtém o diretório de Downloads do usuário
        downloads_path = os.path.expanduser("~/Downloads")

        # Cria o caminho completo do arquivo PDF
        pdf_path = os.path.join(downloads_path, pdf_filename)

        # Cria um arquivo PDF
        c = canvas.Canvas(pdf_path, pagesize=letter)

        # Adiciona o conteúdo do relatório ao PDF
        c.setFont("Helvetica", 12)
        width, height = letter
        lines = relatorio.split("\n")
        y_position = height - 50

        for line in lines:
            # Alinhar à esquerda
            x_position = 50
            c.drawString(x_position, y_position, line)
            y_position -= 15

        # Fecha o arquivo PDF
        c.save()

        return pdf_path

if __name__ == "__main__":
    root = Tk()
    app = CarteiraApp(root)
    root.mainloop()