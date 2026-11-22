import os
import sys
import pandas as pd
import json
import urllib.request
import math

# Função matemática de Poisson
def poisson_probability(lmbda, k):
    return (math.exp(-lmbda) * (lmbda ** k)) / math.factorial(k)

def calcular_poisson_1x2_e_resultados(media_marcar_casa, media_sofrer_casa, media_marcar_fora, media_sofrer_fora):
    esperados_casa = (media_marcar_casa + media_sofrer_fora) / 2.0
    esperados_fora = (media_marcar_fora + media_sofrer_casa) / 2.0

    max_golos = 6
    prob_casa = 0.0
    prob_empate = 0.0
    prob_fora = 0.0
    prob_over25 = 0.0
    prob_btts = 0.0
    resultados_lista = []

    for g_casa in range(max_golos):
        p_c = poisson_probability(esperados_casa, g_casa)
        for g_fora in range(max_golos):
            p_f = poisson_probability(esperados_fora, g_fora)
            prob_combinada = p_c * p_f
            
            resultados_lista.append(((g_casa, g_fora), prob_combinada))
            
            if g_casa > g_fora:
                prob_casa += prob_combinada
            elif g_casa == g_fora:
                prob_empate += prob_combinada
            else:
                prob_fora += prob_combinada
                
            if (g_casa + g_fora) > 2:
                prob_over25 += prob_combinada
                
            if g_casa > 0 and g_fora > 0:
                prob_btts += prob_combinada

    total_1x2 = prob_casa + prob_empate + prob_fora
    prob_casa = (prob_casa / total_1x2) * 100
    prob_empate = (prob_empate / total_1x2) * 100
    prob_fora = (prob_fora / total_1x2) * 100
    prob_over25 = prob_over25 * 100
    prob_btts = prob_btts * 100

    resultados_lista.sort(key=lambda x: x[1], reverse=True)
    top_3_resultados = resultados_lista[:3]
    
    resultados_formatados = []
    for (g_c, g_f), prob in top_3_resultados:
        resultados_formatados.append({
            "placar": f"{g_c} - {g_f}",
            "probabilidade": round(prob * 100, 1)
        })

    return {
        "casa": round(prob_casa, 1),
        "empate": round(prob_empate, 1),
        "fora": round(prob_fora, 1),
        "over25": round(prob_over25, 1),
        "btts": round(prob_btts, 1),
        "resultados_provaveis": resultados_formatados
    }

def selecionar_liga_e_equipa(excel_file, folhas, papel="CASA"):
    print(f"\n\033[93mSelecione a liga para a equipa da {papel}:\033[0m")
    for i, nome_folha in enumerate(folhas):
        print(f"  [{i}] {nome_folha}")
    
    escolha = input(f"Escolha o número ou escreva o nome da liga ({papel}): ").strip()
    try:
        if escolha.isdigit():
            sheet_name = folhas[int(escolha)]
        else:
            sheet_name = escolha if escolha in folhas else folhas[0]
    except Exception:
        sheet_name = folhas[0]

    print(f"\033[94mA carregar liga {sheet_name}...\033[0m")
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    todas_equipas = sorted(list(set(df['HomeTeam'].dropna().unique()) | set(df['AwayTeam'].dropna().unique())))
    
    print(f"\n\033[92mSelecione a equipa da {papel} em {sheet_name}:\033[0m")
    for i, eq in enumerate(todas_equipas):
        print(f"  [{i}] {eq}")

    while True:
        input_escolha = input(f"\nEscolha o número correspondente à equipa da {papel}: ").strip()
        if input_escolha.isdigit():
            idx = int(input_escolha)
            if 0 <= idx < len(todas_equipas):
                equipa_exacta = todas_equipas[idx]
                break
        print("\033[91mOpção inválida. Por favor, tente de novo.\033[0m")
    
    return df, equipa_exacta, sheet_name

def analyze_excel_by_sheets(file_path, model_name="analista-csv"):
    if not os.path.exists(file_path):
        print(f"\033[91mErro: O ficheiro '{file_path}' não foi encontrado.\033[0m")
        return

    try:
        excel_file = pd.ExcelFile(file_path)
        folhas = excel_file.sheet_names
    except Exception as e:
        print(f"\033[91mErro ao abrir o ficheiro Excel: {e}\033[0m")
        return

    # 1. Selecionar Liga e Equipa da CASA
    df_casa_completo, equipa_casa, liga_casa = selecionar_liga_e_equipa(excel_file, folhas, "CASA")

    # 2. Selecionar Liga e Equipa de FORA
    df_fora_completo, equipa_fora, liga_fora = selecionar_liga_e_equipa(excel_file, folhas, "FORA")

    print(f"\n\033[92mAnálise interligas confirmada: {equipa_casa} ({liga_casa} - Casa) vs {equipa_fora} ({liga_fora} - Fora)\033[0m")

    # 3. Filtrar os últimos 10 jogos de cada uma na sua respetiva condição
    jogos_casa = df_casa_completo[df_casa_completo['HomeTeam'] == equipa_casa].tail(10)
    jogos_fora = df_fora_completo[df_fora_completo['AwayTeam'] == equipa_fora].tail(10)

    # Confronto direto (H2H) apenas se forem da mesma liga
    h2h_texto = "  * Equipas de ligas diferentes. Sem histórico recente de confronto direto nos dados nacionais."
    if liga_casa == liga_fora:
        h2h = df_casa_completo[
            ((df_casa_completo['HomeTeam'] == equipa_casa) & (df_casa_completo['AwayTeam'] == equipa_fora)) |
            ((df_casa_completo['HomeTeam'] == equipa_fora) & (df_casa_completo['AwayTeam'] == equipa_casa))
        ].tail(5)
        h2h_lista = []
        for _, row in h2h.iterrows():
            h2h_lista.append(f"  * {row['Date']}: {row['HomeTeam']} {int(row['FTHG'])} - {int(row['FTAG'])} {row['AwayTeam']} (Vencedor: {row['FTR']})")
        if h2h_lista:
            h2h_texto = "\n".join(h2h_lista)

    if len(jogos_casa) == 0 or len(jogos_fora) == 0:
        print("\033[91mErro: Dados insuficientes para as equipas selecionadas.\033[0m")
        return

    # Calcular Médias Reais
    media_m_casa = round(jogos_casa['FTHG'].mean(), 2)
    media_s_casa = round(jogos_casa['FTAG'].mean(), 2)
    media_m_fora = round(jogos_fora['FTAG'].mean(), 2)
    media_s_fora = round(jogos_fora['FTHG'].mean(), 2)

    # Previsão de Poisson
    previsao = calcular_poisson_1x2_e_resultados(media_m_casa, media_s_casa, media_m_fora, media_s_fora)

    # Estatísticas de registos reais
    v_casa = len(jogos_casa[jogos_casa['FTR'] == 'H'])
    e_casa = len(jogos_casa[jogos_casa['FTR'] == 'D'])
    d_casa = len(jogos_casa[jogos_casa['FTR'] == 'A'])

    v_fora = len(jogos_fora[jogos_fora['FTR'] == 'A'])
    e_fora = len(jogos_fora[jogos_fora['FTR'] == 'D'])
    d_fora = len(jogos_fora[jogos_fora['FTR'] == 'H'])

    # Determinar a tendência lógica de golos para injetar no prompt
    tendencia_golos = "MAIS GOLOS (OVER 2.5)" if previsao['over25'] >= 50.0 else "POUCOS GOLOS (UNDER 2.5)"
    btts_tendencia = "SIM (Ambas as equipas têm forte tendência a marcar)" if previsao['btts'] >= 50.0 else "NÃO (Tendência para pelo menos uma das equipas não marcar)"

    # Dupla Possibilidade Calculada de Forma Segura
    dupla_1x = round(previsao['casa'] + previsao['empate'], 1)
    dupla_x2 = round(previsao['fora'] + previsao['empate'], 1)

    # Tabelas Prontas e Perfeitamente Formatadas em Markdown
    tabela_1x2 = (
        f"### 📊 TABELA 1: PROBABILIDADES GLOBAIS (MODELO POISSON INTERLIGAS)\n\n"
        f"| Confronto | Casa (1) | Empate (X) | Fora (2) | Over 2.5 | Ambas Marcam (BTTS) |\n"
        f"| :--- | :---: | :---: | :---: | :---: | :---: |\n"
        f"| **{equipa_casa} ({liga_casa}) vs {equipa_fora} ({liga_fora})** | **{previsao['casa']}%** | **{previsao['empate']}%** | **{previsao['fora']}%** | **{previsao['over25']}%** | **{previsao['btts']}%** |\n"
    )

    tabela_resultados = (
        f"### 🎯 TABELA 2: 3 RESULTADOS EXATOS MAIS PROVÁVEIS\n\n"
        f"| Posição | Placar Estimado | Probabilidade |\n"
        f"| :---: | :---: | :---: |\n"
        f"| **1.º Lugar** | {previsao['resultados_provaveis'][0]['placar']} | **{previsao['resultados_provaveis'][0]['probabilidade']}%** |\n"
        f"| **2.º Lugar** | {previsao['resultados_provaveis'][1]['placar']} | **{previsao['resultados_provaveis'][1]['probabilidade']}%** |\n"
        f"| **3.º Lugar** | {previsao['resultados_provaveis'][2]['placar']} | **{previsao['resultados_provaveis'][2]['probabilidade']}%** |\n"
    )

    factos_ia = f"""
=== DADOS REAIS E VERIFICADOS DO MODELO ===
- EQUIPA DA CASA: {equipa_casa} (Liga: {liga_casa}) -> Últimos 10 jogos em casa: {v_casa} Vitórias, {e_casa} Empates, {d_casa} Derrotas.
  * Média de golos marcados pelo {equipa_casa} em casa: {media_m_casa} por jogo.
  * Média de golos sofridos pelo {equipa_casa} em casa: {media_s_casa} por jogo.

- EQUIPA DE FORA: {equipa_fora} (Liga: {liga_fora}) -> Últimos 10 jogos fora: {v_fora} Vitórias, {e_fora} Empates, {d_fora} Derrotas.
  * Média de golos marcados pelo {equipa_fora} fora: {media_m_fora} por jogo.
  * Média de golos sofridos pelo {equipa_fora} fora: {media_s_fora} por jogo.

- PROBABILIDADES POISSON CALCULADAS PELO SISTEMA (NUNCA ALTERAR ESTES NÚMEROS):
  * Vitória de {equipa_casa} (1): {previsao['casa']}%
  * Empate (X): {previsao['empate']}%
  * Vitória de {equipa_fora} (2): {previsao['fora']}%
  * Dupla Possibilidade {equipa_casa} ou Empate (1X): {dupla_1x}%
  * Dupla Possibilidade {equipa_fora} ou Empate (X2): {dupla_x2}%
  * Over 2.5 (Mais de 2.5 golos no total do jogo): {previsao['over25']}% -> TENDÊNCIA DE GOLOS DO ENCONTRO: {tendencia_golos}
  * Ambas Marcam (BTTS): {previsao['btts']}% -> TENDÊNCIA BTTS: {btts_tendencia}
  * Resultados exatos mais prováveis: 1º: {previsao['resultados_provaveis'][0]['placar']} ({previsao['resultados_provaveis'][0]['probabilidade']}%), 2º: {previsao['resultados_provaveis'][1]['placar']} ({previsao['resultados_provaveis'][1]['probabilidade']}%), 3º: {previsao['resultados_provaveis'][2]['placar']} ({previsao['resultados_provaveis'][2]['probabilidade']}%)
"""

    prompt = (
        "Atua estritamente como o melhor Analista Estatístico e Comentador de Futebol de Portugal.\n"
        "A tua tarefa é escrever uma análise inteligente, curta e direta, baseando-te EXCLUSIVAMENTE nos factos matemáticos fornecidos.\n\n"
        f"{factos_ia}\n\n"
        "REGRAS CRÍTICAS DE COERÊNCIA E LÓGICA (PROIBIDO FALHAR):\n"
        f"1. Se a probabilidade de Over 2.5 é de {previsao['over25']}% (menor que 50%), a tendência é de poucos golos (Under 2.5). Nunca sugiras apostas em Over 2.5 ou em ambas marcam se as probabilidades forem baixas.\n"
        f"2. Se a probabilidade de Dupla Possibilidade (1X) a favor do favorito ({equipa_casa}) é de {dupla_1x}%, isto representa uma segurança gigantesca. Caso queiras sugerir a vitória do favorito mas queiras proteger o empate (já que resultados como 1-1 ou 0-0 são prováveis), recomenda explicitamente 'Dupla Possibilidade ({equipa_casa} ou Empate)' ou 'Empate Anula Aposta (DNB {equipa_casa})'.\n"
        "3. Nunca digas que 'o jogo não tende a ser um empate' se o empate ou placares como 1-1 e 0-0 estão no topo dos resultados prováveis. O empate é uma forte cobertura de segurança.\n"
        "4. Nunca dividas as médias de golos por 10 e não inventes termos incorretos.\n\n"
        "REGRAS DE IDIOMA (PORTUGUÊS DE PORTUGAL PERFEITO):\n"
        "- É terminantemente proibido usar vocabulário brasileiro. Nunca uses: equipe, equipes, time, times, gol, gols, ou a expressão 'marca golos' (usa 'marcar golos' ou 'marcarem golos').\n"
        "- Escreve sempre 'uma média' (feminino) e nunca 'um média'.\n"
        "- Usa sempre: equipa, equipas, golos, relvado, favoritismo, prognóstico.\n\n"
        "ESTRUTURA OBRIGATÓRIA DA RESPOSTA (Escreve apenas estas 3 secções):\n"
        "**ANÁLISE DE FORMA E FORÇA**\n"
        "(Analisa o momento das duas equipas nos respetivos campeonatos com base no registo de vitórias/derrotas fornecido. Máximo 3 frases.)\n\n"
        "**INTERPRETAÇÃO DO MERCADO DE GOLOS**\n"
        "(Compara as médias reais de golos marcados/sofridos com as probabilidades de Over 2.5 e BTTS calculadas pelo sistema. Explica se o jogo tende a ser fechado ou aberto. Máximo 3 frases.)\n\n"
        "**PROGNÓSTICO FUNDAMENTADO**\n"
        "(Indica de forma clara qual é a aposta estatisticamente mais inteligente (ex: Dupla Possibilidade 1X ou Empate Anula Aposta a favor do favorito como segurança de ferro, ou Under 2.5 devido à tendência de poucos golos). Justifica em 2 frases.)"
    )

    url = "http://localhost:11434/api/generate"
    payload = {
        "model": model_name,
        "prompt": prompt,
        "stream": False
    }
    
    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers={'Content-Type': 'application/json'})

    print(f"\033[94mA processar análise interligas perfeita no Ollama ({model_name})...\033[0m")
    try:
        with urllib.request.urlopen(req) as response:
            res = json.loads(response.read().decode('utf-8'))
            print(f"\n\033[92m=== ANÁLISE REAL E INCONTESTÁVEL ({equipa_casa} vs {equipa_fora}) ===\033[0m\n")
            
            # Imprimir as tabelas de Poisson limpas e alinhadas
            print(tabela_1x2)
            print(tabela_resultados)
            
            # Imprimir a análise inteligente da IA
            print("### 📝 ANÁLISE DO MODELO DE INTELIGÊNCIA ARTIFICIAL\n")
            print(res.get('response', 'Sem resposta gerada.'))
            print("\n\033[92m============================================================\033[0m\n")
    except Exception as e:
        print(f"\033[91mErro ao conectar ao Ollama: {e}\033[0m")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso correto: python analisar_dados.py <caminho_do_ficheiro.xlsx> [modelo]")
    else:
        file = sys.argv[1]
        model = sys.argv[2] if len(sys.argv) > 2 else "analista-csv"
        analyze_excel_by_sheets(file, model)
