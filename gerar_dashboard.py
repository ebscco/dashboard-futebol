import pandas as pd
import json
import os

excel_path = 'all-euro-data-2025-2026.xlsx'
xl = pd.ExcelFile(excel_path)

leagues_info = {
    'P1': 'Portugal (Liga Portugal)',
    'E0': 'Inglaterra (Premier League)',
    'SP1': 'Espanha (La Liga)',
    'I1': 'Itália (Serie A)',
    'D1': 'Alemanha (Bundesliga)',
    'F1': 'França (Ligue 1)'
}

dashboard_data = {}

for code, name in leagues_info.items():
    if code in xl.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=code)
        # Limpar dados NaN em colunas essenciais
        df = df.dropna(subset=['HomeTeam', 'AwayTeam', 'FTHG', 'FTAG'])
        
        # Calcular golos por equipa
        home_g = df.groupby('HomeTeam')['FTHG'].sum()
        away_g = df.groupby('AwayTeam')['FTAG'].sum()
        total_g = home_g.add(away_g, fill_value=0)
        
        home_p = df.groupby('HomeTeam')['FTHG'].count()
        away_p = df.groupby('AwayTeam')['FTAG'].count()
        total_p = home_p.add(away_p, fill_value=0)
        
        stats = []
        for team in total_g.index:
            hg = home_g.get(team, 0)
            ag = away_g.get(team, 0)
            tg = total_g[team]
            gp = total_p[team]
            stats.append({
                "team": team,
                "matches": int(gp),
                "goals": int(tg),
                "avg": round(tg / gp, 2) if gp > 0 else 0
            })
        
        stats.sort(key=lambda x: x['goals'], reverse=True)
        
        # Últimos jogos
        recent = df.tail(15).copy()
        matches_list = []
        for _, row in recent.iterrows():
            date_str = str(row['Date'])[:10] if 'Date' in row else ''
            matches_list.append({
                "date": date_str,
                "home": row['HomeTeam'],
                "away": row['AwayTeam'],
                "fthg": int(row['FTHG']),
                "ftag": int(row['FTAG']),
                "ftr": row['FTR']
            })
            
        dashboard_data[code] = {
            "name": name,
            "total_matches": len(df),
            "total_goals": int(df['FTHG'].sum() + df['FTAG'].sum()),
            "avg_goals": round((df['FTHG'].sum() + df['FTAG'].sum()) / len(df), 2) if len(df) > 0 else 0,
            "teams": stats,
            "matches": matches_list[::-1]
        }

html_content = f"""<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Futebol Europeu 2025/2026 - JARVIS</title>
    <style>
        :root {{
            --bg-color: #0f172a;
            --card-bg: #1e293b;
            --card-border: #334155;
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --accent: #38bdf8;
            --accent-gradient: linear-gradient(135deg, #38bdf8 0%, #818cf8 100%);
            --success: #22c55e;
            --warning: #f59e0b;
        }}
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
        }}
        body {{
            background-color: var(--bg-color);
            color: var(--text-primary);
            min-height: 100vh;
            padding: 20px;
        }}
        header {{
            text-align: center;
            margin-bottom: 30px;
            padding: 20px;
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 16px;
            box-shadow: 0 10px 25px -5px rgba(0,0,0,0.3);
        }}
        header h1 {{
            font-size: 2rem;
            background: var(--accent-gradient);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 8px;
        }}
        header p {{
            color: var(--text-secondary);
            font-size: 0.95rem;
        }}
        .league-tabs {{
            display: flex;
            justify-content: center;
            gap: 10px;
            flex-wrap: wrap;
            margin-bottom: 30px;
        }}
        .tab-btn {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            color: var(--text-secondary);
            padding: 10px 18px;
            border-radius: 10px;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.2s ease;
        }}
        .tab-btn:hover {{
            border-color: var(--accent);
            color: var(--text-primary);
        }}
        .tab-btn.active {{
            background: var(--accent-gradient);
            color: #0f172a;
            border-color: transparent;
            box-shadow: 0 4px 12px rgba(56, 189, 248, 0.3);
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
        }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .stat-card {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 14px;
            padding: 20px;
            text-align: center;
        }}
        .stat-card h3 {{
            color: var(--text-secondary);
            font-size: 0.85rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 8px;
        }}
        .stat-card .value {{
            font-size: 2rem;
            font-weight: 700;
            color: var(--text-primary);
        }}
        .content-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }}
        @media (max-width: 900px) {{
            .content-grid {{
                grid-template-columns: 1fr;
            }}
        }}
        .panel {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 14px;
            padding: 20px;
            overflow: hidden;
        }}
        .panel h2 {{
            font-size: 1.2rem;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 8px;
            color: var(--accent);
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 0.9rem;
        }}
        th, td {{
            padding: 10px 12px;
            border-bottom: 1px solid var(--card-border);
        }}
        th {{
            color: var(--text-secondary);
            font-weight: 600;
            background: rgba(255,255,255,0.02);
        }}
        tr:hover td {{
            background: rgba(255,255,255,0.03);
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            color: var(--text-secondary);
            font-size: 0.85rem;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>⚽ Dashboard de Futebol Europeu</h1>
            <p>Estatísticas atualizadas da época 2025/2026 | JARVIS Agent</p>
        </header>

        <div class="league-tabs" id="leagueTabs"></div>

        <div class="stats-grid">
            <div class="stat-card">
                <h3>Total de Jogos</h3>
                <div class="value" id="totalMatches">-</div>
            </div>
            <div class="stat-card">
                <h3>Total de Golos</h3>
                <div class="value" id="totalGoals">-</div>
            </div>
            <div class="stat-card">
                <h3>Média de Golos / Jogo</h3>
                <div class="value" id="avgGoals">-</div>
            </div>
        </div>

        <div class="content-grid">
            <div class="panel">
                <h2>🏆 Equipas Mais Goleadoras</h2>
                <div style="max-height: 400px; overflow-y: auto;">
                    <table>
                        <thead>
                            <tr>
                                <th>Equipa</th>
                                <th style="text-align: center;">Jogos</th>
                                <th style="text-align: center;">Golos</th>
                                <th style="text-align: center;">Média</th>
                            </tr>
                        </thead>
                        <tbody id="teamsTable"></tbody>
                    </table>
                </div>
            </div>

            <div class="panel">
                <h2>🕒 Últimos Jogos</h2>
                <div style="max-height: 400px; overflow-y: auto;">
                    <table>
                        <thead>
                            <tr>
                                <th>Data</th>
                                <th>Jogo</th>
                                <th style="text-align: center;">Resultado</th>
                            </tr>
                        </thead>
                        <tbody id="matchesTable"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <div class="footer">
            Criado pelo JARVIS 🤖 | Dados do Excel `all-euro-data-2025-2026.xlsx`
        </div>
    </div>

    <script>
        const data = {json.dumps(dashboard_data, ensure_ascii=False)};
        
        const tabsContainer = document.getElementById('leagueTabs');
        const totalMatchesEl = document.getElementById('totalMatches');
        const totalGoalsEl = document.getElementById('totalGoals');
        const avgGoalsEl = document.getElementById('avgGoals');
        const teamsTableEl = document.getElementById('teamsTable');
        const matchesTableEl = document.getElementById('matchesTable');

        let currentLeague = Object.keys(data)[0];

        function renderTabs() {{
            tabsContainer.innerHTML = '';
            for (const [code, info] of Object.entries(data)) {{
                const btn = document.createElement('button');
                btn.className = `tab-btn ${{code === currentLeague ? 'active' : ''}}`;
                btn.innerText = info.name;
                btn.onclick = () => {{
                    currentLeague = code;
                    renderTabs();
                    renderDashboard();
                }};
                tabsContainer.appendChild(btn);
            }}
        }}

        function renderDashboard() {{
            const league = data[currentLeague];
            totalMatchesEl.innerText = league.total_matches;
            totalGoalsEl.innerText = league.total_goals;
            avgGoalsEl.innerText = league.avg_goals;

            // Render Teams
            teamsTableEl.innerHTML = '';
            league.teams.forEach(t => {{
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td><strong>${{t.team}}</strong></td>
                    <td style="text-align: center;">${{t.matches}}</td>
                    <td style="text-align: center; color: #38bdf8; font-weight: 600;">${{t.goals}}</td>
                    <td style="text-align: center;">${{t.avg}}</td>
                `;
                teamsTableEl.appendChild(tr);
            }});

            // Render Matches
            matchesTableEl.innerHTML = '';
            league.matches.forEach(m => {{
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="color: #94a3b8; font-size: 0.85rem;">${{m.date}}</td>
                    <td>${{m.home}} vs ${{m.away}}</td>
                    <td style="text-align: center; font-weight: 600;">${{m.fthg}} - ${{m.ftag}}</td>
                `;
                matchesTableEl.appendChild(tr);
            }});
        }}

        renderTabs();
        renderDashboard();
    </script>
</body>
</html>
"""

os.makedirs('analista', exist_ok=True)
with open('analista/dashboard.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("Dashboard gerado com sucesso em analista/dashboard.html!")
python_code = os.path.abspath('analista/dashboard.html')
print(f"Caminho: {python_code}")
