def to_excel_package(div, stats, meta=None):
    """
    Gera XLSX com:
      - Resumo (formatado)
      - Divergencias (formatado)
      - Tratativa (formatado)
    """
    meta = meta or {}
    out = BytesIO()

    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb = w.book

        fmt_title = wb.add_format({"bold": True, "font_size": 16})
        fmt_label = wb.add_format({"bold": True})
        fmt_hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter"})
        fmt_money = wb.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_text = wb.add_format({"border": 1})
        fmt_date = wb.add_format({"num_format": "dd/mm/yyyy", "border": 1})

        # ---------------- RESUMO ----------------
        resumo_rows = [
            ["Saldo anterior (antes do 1º movimento) - Financeiro", stats.get("saldo_ant_fin", np.nan)],
            ["Saldo anterior (antes do 1º movimento) - Contábil", stats.get("saldo_ant_led", np.nan)],
            ["Diferença de saldo anterior (Fin - Cont)", stats.get("diff_saldo_ant", np.nan)],
            [None, None],
            ["Saldo final (último movimento) - Financeiro", stats.get("saldo_fin", np.nan)],
            ["Saldo final (último movimento) - Contábil", stats.get("saldo_led", np.nan)],
            ["Diferença final (Fin - Cont)", stats.get("diff_final", np.nan)],
            [None, None],
            ["Soma dos movimentos só no Financeiro (aba Divergencias)", stats.get("fin_pend_val", np.nan)],
            ["Soma dos movimentos só no Contábil (aba Divergencias)", stats.get("led_pend_val", np.nan)],
            ["Impacto líquido dos movimentos (Fin - Cont)", stats.get("impacto", np.nan)],
            ["Diferença esperada (Dif. saldo anterior + Impacto)", stats.get("diff_esperada", np.nan)],
            ["Conferência (Dif. final - Dif. esperada) -> precisa zerar", stats.get("conferencia", np.nan)],
        ]
        resumo = pd.DataFrame(resumo_rows, columns=["Metrica", "Valor"])
        resumo.to_excel(w, index=False, sheet_name="Resumo", startrow=2)

        ws = w.sheets["Resumo"]
        ws.write(0, 0, "ConciliaMais — Resumo da Conciliação", fmt_title)
        ws.write(1, 0, f"Processado em: {meta.get('generated_at','')}", None)

        ws.set_row(2, 20, fmt_hdr)
        ws.set_column(0, 0, 62)
        ws.set_column(1, 1, 22, fmt_money)

        # ---------------- DIVERGENCIAS ----------------
        div_out = div.copy()

        # limpa 'nan' textual
        for c in div_out.columns:
            div_out[c] = div_out[c].replace(["nan", "None"], "", regex=False)

        div_out.to_excel(w, index=False, sheet_name="Divergencias", startrow=6)
        ws = w.sheets["Divergencias"]

        ws.write(0, 0, "ConciliaMais — Divergências (Excel igual à tela)", fmt_title)
        ws.write(1, 0, f"Processado em: {meta.get('generated_at','')}", None)

        # bloco “filtros/totalizadores” (opcional)
        ws.write(3, 0, "Origem:", fmt_label);        ws.write(3, 1, meta.get("origem",""))
        ws.write(4, 0, "Visualização:", fmt_label);  ws.write(4, 1, meta.get("visualizacao",""))
        ws.write(5, 0, "Busca:", fmt_label);         ws.write(5, 1, meta.get("busca",""))

        ws.write(3, 4, "Total do filtro:", fmt_label);  ws.write(3, 5, meta.get("total_filtro", 0.0), fmt_money)
        ws.write(4, 4, "Total em aberto:", fmt_label);  ws.write(4, 5, meta.get("total_aberto", 0.0), fmt_money)

        # header da tabela
        ws.set_row(6, 20, fmt_hdr)
        ws.freeze_panes(7, 0)

        # colunas (ajuste fino)
        col_map = {name: i for i, name in enumerate(div_out.columns)}
        # widths padrão
        ws.set_column(0, 0, 6)   # ID se existir
        for c, idx in col_map.items():
            if c.upper() in ["ORIGEM"]:
                ws.set_column(idx, idx, 18)
            elif c.upper() in ["DATA"]:
                ws.set_column(idx, idx, 12, fmt_date)
            elif c.upper() in ["DOCUMENTO"]:
                ws.set_column(idx, idx, 18)
            elif c.upper() in ["HISTORICO_OPERACAO"]:
                ws.set_column(idx, idx, 52)
            elif c.upper() in ["CHAVE_DOC"]:
                ws.set_column(idx, idx, 18)
            elif c.upper() in ["VALOR"]:
                ws.set_column(idx, idx, 16, fmt_money)
            else:
                ws.set_column(idx, idx, 20)

        # aplica borda nas células (opcional: “limpinho”)
        # (xlsxwriter não aplica fácil em bloco sem loop; então deixei por coluna principal)

        # ---------------- TRATATIVA ----------------
        trat = build_tratativa(div)
        trat.to_excel(w, index=False, sheet_name="Tratativa")
        ws = w.sheets["Tratativa"]
        ws.freeze_panes(1, 0)
        ws.set_row(0, 20, fmt_hdr)
        ws.set_column(0, 0, 18)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 3, 48)
        ws.set_column(4, 4, 16, fmt_money)
        ws.set_column(5, 9, 22)

    out.seek(0)
    return out.getvalue()
