pip install streamlit sqlalchemy pandas openpyxl matplotlib seaborn
# -*- coding: utf-8 -*-
import os
import sys
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sqlalchemy import create_engine, Column, Integer, String, Boolean, DateTime
from sqlalchemy.orm import declarative_base, sessionmaker

# Configuração do Banco de Dados SQLite com SQLAlchemy
Base = declarative_base()

class NaoConformidadeModel(Base):
    __tablename__ = 'nao_conformidades'
    
    id_nc = Column(String, primary_key=True)
    data_identificacao = Column(DateTime, default=datetime.now)
    descricao_saida = Column(String, nullable=False)
    processo_origem = Column(String, nullable=False)
    responsavel_identificacao = Column(String, nullable=False)
    impacto_potencial = Column(String, nullable=False)  # Baixo, Médio, Alto
    status = Column(String, default="Identificado/Segregado")
    tratativa_adotada = Column(String, nullable=True)
    detalhes_tratativa = Column(String, nullable=True)
    autorizado_por = Column(String, nullable=True)
    data_fechamento = Column(DateTime, nullable=True)
    informacao_documentada_retida = Column(Boolean, default=False)
    prazo_sla = Column(DateTime, nullable=False)

class GestaoNaoConformidadeISO9001:
    """
    Classe de gestão de saídas não conformes com banco SQLite, SLAs e Exportações.
    """
    def __init__(self, db_path: str = "iso9001_qualidade.db"):
        self.engine = create_engine(f"sqlite:///{db_path}", echo=False)
        Base.metadata.create_all(self.engine)
        self.Session = sessionmaker(bind=self.engine)

    def _gerar_id_unico(self, session) -> str:
        ano_atual = datetime.now().strftime("%Y")
        count = session.query(NaoConformidadeModel).filter(NaoConformidadeModel.id_nc.like(f"NC-{ano_atual}-%")).count()
        return f"NC-{ano_atual}-{(count + 1):04d}"

    def cadastrar_saida_nao_conforme(self, descricao_saida: str, processo_origem: str, responsavel_identificacao: str, impacto_potencial: str, dias_sla: int = 15, data_retroativa=None) -> str:
        session = self.Session()
        id_nc = self._gerar_id_unico(session)
        
        data_id = data_retroativa if data_retroativa else datetime.now()
        prazo = data_id + timedelta(days=dias_sla)
        
        nova_nc = NaoConformidadeModel(
            id_nc=id_nc,
            data_identificacao=data_id,
            descricao_saida=descricao_saida,
            processo_origem=processo_origem,
            responsavel_identificacao=responsavel_identificacao,
            impacto_potencial=impacto_potencial,
            status="Identificado/Segregado",
            prazo_sla=prazo
        )
        session.add(nova_nc)
        session.commit()
        session.close()
        return id_nc

    def aplicar_tratativa(self, id_nc: str, acao_tratativa: str, detalhes: str) -> None:
        acoes_validas = ['Correção', 'Segregação/Retenção', 'Informação ao Cliente', 'Concessão']
        if acao_tratativa not in acoes_validas:
            raise ValueError(f"Ação deve ser uma das seguintes: {acoes_validas}")

        session = self.Session()
        nc = session.query(NaoConformidadeModel).filter_by(id_nc=id_nc).first()
        if not nc:
            session.close()
            raise ValueError(f"Registro {id_nc} não encontrado.")
        
        nc.tratativa_adotada = acao_tratativa
        nc.detalhes_tratativa = detalhes
        nc.status = "Em Tratativa"
        session.commit()
        session.close()

    def fechar_e_autorizar(self, id_nc: str, responsavel_autorizacao: str, data_retroativa_fechamento=None) -> None:
        session = self.Session()
        nc = session.query(NaoConformidadeModel).filter_by(id_nc=id_nc).first()
        if not nc:
            session.close()
            raise ValueError(f"Registro {id_nc} não encontrado.")
        if not nc.tratativa_adotada:
            session.close()
            raise PermissionError("Não é possível fechar sem uma tratativa aplicada.")

        nc.autorizado_por = responsavel_autorizacao
        nc.data_fechamento = data_retroativa_fechamento if data_retroativa_fechamento else datetime.now()
        nc.status = "Concluído"
        nc.informacao_documentada_retida = True
        session.commit()
        session.close()

    def extrair_dataframe(self) -> pd.DataFrame:
        session = self.Session()
        query = session.query(NaoConformidadeModel)
        df = pd.read_sql(query.statement, session.bind)
        session.close()
        
        # Garante a tipagem de data para evitar problemas
        if not df.empty:
            df['data_identificacao'] = pd.to_datetime(df['data_identificacao'])
            df['prazo_sla'] = pd.to_datetime(df['prazo_sla'])
            df['data_fechamento'] = pd.to_datetime(df['data_fechamento'])
        return df

    def exportar_excel(self, filename: str = "relatorio_iso9001.xlsx") -> str:
        df = self.extrair_dataframe()
        if df.empty:
            return "Nenhum registro encontrado para exportação."
        
        # Calcula SLA no DataFrame
        status_sla = []
        agora = datetime.now()
        for idx, row in df.iterrows():
            if row['status'] == 'Concluído':
                if row['data_fechamento'] <= row['prazo_sla']:
                    status_sla.append('Fechada no Prazo')
                else:
                    status_sla.append('Fechada com Atraso')
            else:
                if agora > row['prazo_sla']:
                    status_sla.append('Atrasada (Aberta)')
                else:
                    status_sla.append('No Prazo (Aberta)')
        
        df['Indicador_SLA'] = status_sla
        df.to_excel(filename, index=False)
        return filename

# Popula dados de simulação iniciais se o banco estiver vazio
def popular_dados_teste(sistema):
    df = sistema.extrair_dataframe()
    if not df.empty:
        return
    
    processos = ["Produção", "Logística", "Engenharia", "Suprimentos", "Qualidade"]
    impactos = ["Baixo", "Médio", "Alto"]
    tratativas = ['Correção', 'Segregação/Retenção', 'Informação ao Cliente', 'Concessão']
    
    for i in range(30):
        proc = random.choice(processos)
        imp = random.choices(impactos, weights=[0.5, 0.3, 0.2])[0]
        dias_atras = random.randint(5, 25)
        data_id = datetime.now() - timedelta(days=dias_atras)
        
        # SLA padrão de 15 dias
        id_nc = sistema.cadastrar_saida_nao_conforme(
            descricao_saida=f"Desvio operacional simulado #{i}",
            processo_origem=proc,
            responsavel_identificacao="Sistema Inteligente",
            impacto_potencial=imp,
            dias_sla=15,
            data_retroativa=data_id
        )
        
        if random.random() > 0.4:
            trat = random.choice(tratativas)
            sistema.aplicar_tratativa(id_nc, trat, "Tratativa padrão executada.")
            if random.random() > 0.3:
                # Altera data de fechamento para simular dentro ou fora do SLA
                dias_fechar = random.randint(5, 20)
                sistema.fechar_e_autorizar(id_nc, "Gestor Qualidade", data_retroativa_fechamento=data_id + timedelta(days=dias_fechar))

import random

def exibir_menu():
    print("\n" + "="*50)
    print("  SISTEMA DE GESTÃO ISO 9001:2026 - CLI INTERATIVA")
    print("="*50)
    print("[1] Cadastrar Nova Saída Não Conforme")
    print("[2] Aplicar Tratativa/Ação de Contenção")
    print("[3] Autorizar Fechamento (Reter Informação Documentada)")
    print("[4] Listar Ocorrências Ativas e Alertas de SLA")
    print("[5] Renderizar Painel de Indicadores (Gráficos)")
    print("[6] Exportar Relatório Consolidado para Excel")
    print("[0] Sair do Sistema")
    print("="*50)

def main():
    sistema = GestaoNaoConformidadeISO9001()
    popular_dados_teste(sistema)
    
    while True:
        exibir_menu()
        opcao = input("Selecione uma opção: ").strip()
        
        if opcao == "1":
            print("\n--- CADASTRAR NÃO CONFORMIDADE ---")
            desc = input("Descrição da Saída Não Conforme: ")
            proc = input("Processo de Origem (Ex: Produção, Engenharia): ")
            resp = input("Responsável pela Identificação: ")
            imp = input("Impacto Potencial (Baixo, Médio, Alto): ")
            try:
                dias = int(input("Prazo de Resolução SLA (em dias, Padrão 15): ") or 15)
            except ValueError:
                dias = 15
            
            id_nc = sistema.cadastrar_saida_nao_conforme(desc, proc, resp, imp, dias)
            print(f"\n[SUCESSO] Registrado com sucesso! Código gerado: {id_nc}")
            
        elif opcao == "2":
            print("\n--- APLICAR TRATATIVA ---")
            id_nc = input("Informe o Código da NC (Ex: NC-2026-0001): ").strip()
            print("Tipos válidos: Correção, Segregação/Retenção, Informação ao Cliente, Concessão")
            trat = input("Tipo de Tratativa: ").strip()
            detalhes = input("Detalhes técnicos da ação: ")
            try:
                sistema.aplicar_tratativa(id_nc, trat, detalles)
                print(f"\n[SUCESSO] Tratativa aplicada à {id_nc}.")
            except Exception as e:
                print(f"\n[ERRO] Falha ao aplicar tratativa: {e}")
                
        elif opcao == "3":
            print("\n--- FECHAMENTO E AUTORIZAÇÃO ---")
            id_nc = input("Informe o Código da NC: ").strip()
            resp = input("Nome da Autoridade que Autoriza o Fechamento: ")
            try:
                sistema.fechar_e_autorizar(id_nc, resp)
                print(f"\n[SUCESSO] Ocorrência {id_nc} encerrada. Informação documentada armazenada em conformidade.")
            except Exception as e:
                print(f"\n[ERRO] Falha ao encerrar: {e}")
                
        elif opcao == "4":
            print("\n--- LISTAGEM ATIVA COM REGRAS DE SLA ---")
            df = sistema.extrair_dataframe()
            if df.empty:
                print("Nenhum registro em banco.")
                continue
                
            agora = datetime.now()
            print(f"{'ID':<12} | {'Processo':<12} | {'Status':<22} | {'Prazo SLA':<10} | {'Status SLA'}")
            print("-" * 75)
            
            for idx, row in df.iterrows():
                p_sla = row['prazo_sla']
                status = row['status']
                
                if status == 'Concluído':
                    status_sla = "✅ FECHADA NO PRAZO" if row['data_fechamento'] <= p_sla else "⚠️ FECHADA COM ATRASO"
                else:
                    status_sla = "🚨 ATRASADA" if agora > p_sla else "⏳ No Prazo"
                    
                print(f"{row['id_nc']:<12} | {row['processo_origem']:<12} | {status:<22} | {p_sla.strftime('%Y-%m-%d'):<10} | {status_sla}")
                
        elif opcao == "5":
            print("\nGerando e renderizando gráficos de controle...")
            df = sistema.extrair_dataframe()
            if df.empty:
                print("Sem dados suficientes.")
                continue
            
            sns.set_theme(style="whitegrid")
            fig, axes = plt.subplots(1, 2, figsize=(14, 6))
            fig.suptitle("Métricas de Qualidade Cláusula 8.7 / 9.3", fontsize=14, fontweight='bold')
            
            # Grafico 1: Processos
            dados_proc = df['processo_origem'].value_counts()
            axes[0].pie(dados_proc, labels=dados_proc.index, autopct='%1.1f%%', startangle=140, colors=sns.color_palette("viridis", len(dados_proc)))
            axes[0].set_title("Ocorrências por Processo")
            
            # Grafico 2: SLA Status
            agora = datetime.now()
            sla_list = []
            for idx, row in df.iterrows():
                if row['status'] == 'Concluído':
                    sla_list.append("Concluída no Prazo" if row['data_fechamento'] <= row['prazo_sla'] else "Concluída com Atraso")
                else:
                    sla_list.append("Atrasada (Aberta)" if agora > row['prazo_sla'] else "No Prazo (Aberta)")
            df['SLA_Ctx'] = sla_list
            
            sns.countplot(data=df, x='SLA_Ctx', ax=axes[1], palette="coolwarm")
            axes[1].set_title("Status Geral de Cumprimento do SLA")
            axes[1].set_xlabel("Status do Prazo")
            axes[1].set_ylabel("Quantidade")
            
            plt.tight_layout()
            plt.show()
            
        elif opcao == "6":
            filename = sistema.exportar_excel()
            print(f"\n[SUCESSO] Relatório estruturado exportado para: {filename}")
            
        elif opcao == "0":
            print("\nSaindo do sistema de auditoria...")
            break
        else:
            print("\nOpção Inválida. Tente novamente.")

if __name__ == "__main__":
    main()
