#!/usr/bin/env python3
"""
Generate season-by-season baseball card checklists for MLB teams in Lahman's
American League (AL) and National League (NL) records.

Outputs:
    checklists/team-name.pdf  - three printable pages per season

Data source: Lahman Baseball Database, accessed through pylahman.
By default, one PDF is generated for every MLB team in the downloaded
database from 1960 through the latest available season for that team.
Newer teams begin with their first available season. Historical non-MLB
leagues in Lahman are excluded.

Install:
    python -m pip install pylahman pandas openpyxl reportlab

Run:
    python card_checklists.py
    python card_checklists.py --team-id NYA --start 1970 --end 2025
    python card_checklists.py --start 1970 --end 2025
    C:/Users/admin/AppData/Local/Programs/Python/Python314/python.exe card_checklists.py --start 1970 --end 2025
"""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.graphics.shapes import Drawing, Rect
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
)


PDF_FONT = "Helvetica"
PDF_FONT_BOLD = "Helvetica-Bold"
MLB_LEAGUES = {"AL", "NL"}
MIN_POSITION_PLAYER_GAMES = 1
MIN_PITCHER_IP = 20
MIN_PITCHER_STARTS = 5
POSITION_COLUMNS = [
    "G_c", "G_1b", "G_2b", "G_3b", "G_ss", "G_lf", "G_cf", "G_rf",
    "G_of", "G_dh", "G_ph", "G_pr"
]

POSITION_LABELS = {
    "G_c": "C", "G_1b": "1B", "G_2b": "2B", "G_3b": "3B",
    "G_ss": "SS", "G_lf": "LF", "G_cf": "CF", "G_rf": "RF",
    "G_of": "OF", "G_dh": "DH", "G_ph": "PH", "G_pr": "PR",
}


def display_positions(row):
    excluded = {"G_ph", "G_pr"}
    infield = {"G_1b", "G_2b", "G_3b", "G_ss"}
    outfield = {"G_lf", "G_cf", "G_rf", "G_of"}
    played = {column for column in POSITION_COLUMNS if column not in excluded and row[column] > 0}
    labels = []
    if len(played & infield) >= 3:
        labels.append("INF")
        played -= infield
    else:
        labels.extend(
            POSITION_LABELS[column]
            for column in POSITION_COLUMNS
            if column in played and column not in outfield
        )
    if played & outfield:
        labels.append("OF")
    return ", ".join(dict.fromkeys(labels))


def get_lahman_tables():
    try:
        import pylahman
        batting = pylahman.Batting()
        pitching = pylahman.Pitching()
        people = pylahman.People()
        appearances = pylahman.Appearances()
        teams = pylahman.Teams()
        managers = pylahman.Managers()
        return batting, pitching, people, appearances, teams, managers
    except Exception as exc:
        raise RuntimeError(
            "Could not download/read Lahman data through pylahman. "
            "Install or upgrade pylahman and check your internet connection:\n"
            "  python -m pip install --upgrade pylahman"
        ) from exc


def prepare_data(
    batting, pitching, people, appearances, teams, managers,
    team_ids, start, end, team_seasons=None
):
    for df in (batting, pitching, appearances, teams, people, managers):
        df.columns = [str(c) for c in df.columns]

    required = {
        "batting": {"playerID", "yearID", "teamID", "G", "AB", "H", "HR", "RBI"},
        "pitching": {"playerID", "yearID", "teamID", "IPOuts", "G", "GS", "W", "L", "ER", "SO", "SV"},
        "appearances": {"playerID", "yearID", "teamID"},
        "people": {"playerID", "nameFirst", "nameLast"},
    }
    for name, cols in required.items():
        df = {"batting": batting, "pitching": pitching,
              "appearances": appearances, "people": people}[name]
        missing = cols - set(df.columns)
        if missing:
            raise RuntimeError(f"Lahman {name} table is missing columns: {sorted(missing)}")

    bat = batting[(batting["teamID"].isin(team_ids)) &
                  (batting["yearID"].between(start, end))].copy()
    if team_seasons is not None:
        bat = bat[bat["yearID"].isin(team_seasons)]
    for col in ["G", "AB", "H", "2B", "3B", "HR", "RBI", "BB", "HBP", "SF"]:
        if col not in bat.columns:
            bat[col] = 0
        bat[col] = pd.to_numeric(bat[col], errors="coerce").fillna(0)
    # Sum selected-team stints for players traded during a season.
    bat = bat.groupby(["yearID", "playerID"], as_index=False).agg(
        G=("G", "sum"), AB=("AB", "sum"), H=("H", "sum"),
        doubles=("2B", "sum"), triples=("3B", "sum"), HR=("HR", "sum"),
        RBI=("RBI", "sum"), BB=("BB", "sum"), HBP=("HBP", "sum"), SF=("SF", "sum")
    )
    bat["PA"] = bat["AB"] + bat["BB"] + bat["HBP"] + bat["SF"]
    bat["AVG"] = bat.apply(lambda r: r["H"] / r["AB"] if r["AB"] else 0, axis=1)
    bat["OPS"] = bat.apply(
        lambda r: (
            (r["H"] + r["BB"] + r["HBP"]) /
            (r["AB"] + r["BB"] + r["HBP"] + r["SF"])
            + (r["H"] + 2 * r["doubles"] + 3 * r["triples"] + 3 * r["HR"]) / r["AB"]
        ) if r["AB"] and (r["AB"] + r["BB"] + r["HBP"] + r["SF"]) else 0,
        axis=1,
    )

    app = appearances[(appearances["teamID"].isin(team_ids)) &
                       (appearances["yearID"].between(start, end))].copy()
    if team_seasons is not None:
        app = app[app["yearID"].isin(team_seasons)]
    app = app.groupby(["yearID", "playerID"], as_index=False).sum(numeric_only=True)
    for col in POSITION_COLUMNS:
        if col not in app.columns:
            app[col] = 0
        app[col] = pd.to_numeric(app[col], errors="coerce").fillna(0)
    non_pitcher_position_columns = [
        column for column in POSITION_COLUMNS if column not in {"G_ph", "G_pr"}
    ]
    app["non_pitcher_games"] = app[non_pitcher_position_columns].sum(axis=1)
    if "G_p" in app.columns:
        app["G_p"] = pd.to_numeric(app["G_p"], errors="coerce").fillna(0)
    else:
        app["G_p"] = 0
    app["Position"] = app.apply(display_positions, axis=1)

    hitters = bat.merge(app[["yearID", "playerID", "non_pitcher_games", "Position"]],
                        on=["yearID", "playerID"], how="left")
    hitters["non_pitcher_games"] = hitters["non_pitcher_games"].fillna(0)
    position_eligible = hitters["non_pitcher_games"] >= MIN_POSITION_PLAYER_GAMES
    hitters = hitters[position_eligible].copy()
    hitters = hitters.sort_values(
        ["yearID", "G", "PA", "playerID"], ascending=[True, False, False, True]
    )
    hitters["rank"] = hitters.groupby("yearID").cumcount() + 1
    hitters = hitters[hitters["rank"] <= 18].copy()
    hitters["section"] = "Position player"
    hitters["qualifying_stat"] = hitters["G"].astype(int).astype(str) + " G"
    hitters["Common stats"] = [
        f'G{int(row["G"])} AB{int(row["AB"])} H{int(row["H"])} HR{int(row["HR"])} '
        f'RBI{int(row["RBI"])} AVG {row["AVG"]:.3f}'.replace("AVG 0.", "AVG .")
        for _, row in hitters.iterrows()
    ]

    pit = pitching[(pitching["teamID"].isin(team_ids)) &
                   (pitching["yearID"].between(start, end))].copy()
    if team_seasons is not None:
        pit = pit[pit["yearID"].isin(team_seasons)]
    for col in ["IPOuts", "G", "GS", "W", "L", "ER", "SO", "SV"]:
        pit[col] = pd.to_numeric(pit[col], errors="coerce").fillna(0)
    pit = pit.groupby(["yearID", "playerID"], as_index=False).agg(
        IPOuts=("IPOuts", "sum"), G=("G", "sum"), GS=("GS", "sum"),
        W=("W", "sum"), L=("L", "sum"), ER=("ER", "sum"),
        SO=("SO", "sum"), SV=("SV", "sum")
    )
    pitcher_eligible = (
        (pit["IPOuts"] >= MIN_PITCHER_IP * 3) |
        (pit["GS"] >= MIN_PITCHER_STARTS)
    )
    pit = pit[pitcher_eligible].copy()
    pit["IP"] = pit["IPOuts"] / 3.0
    pit["ERA"] = pit.apply(
        lambda r: (r["ER"] * 27 / r["IPOuts"]) if r["IPOuts"] else 0, axis=1
    )
    pit = pit.sort_values(
        ["yearID", "IPOuts", "SO", "playerID"], ascending=[True, False, False, True]
    )
    pit["rank"] = pit.groupby("yearID").cumcount() + 1
    pit = pit[pit["rank"] <= 9].copy()
    pit["section"] = "Pitcher"
    pit["qualifying_stat"] = pit["IPOuts"].apply(format_innings)
    pit["Common stats"] = [
        f'G{int(row["G"])} GS{int(row["GS"])} W-L {int(row["W"])}-{int(row["L"])} '
        f'ERA {row["ERA"]:.2f} SO{int(row["SO"])} SV{int(row["SV"])}'
        for _, row in pit.iterrows()
    ]

    names = people[["playerID", "nameFirst", "nameLast"]].copy()
    names["Player"] = (
        names["nameFirst"].fillna("").astype(str).str.strip() + " " +
        names["nameLast"].fillna("").astype(str).str.strip()
    ).str.strip()
    names = names[["playerID", "Player"]]

    hitters = hitters.merge(names, on="playerID", how="left")
    pit = pit.merge(names, on="playerID", how="left")
    hitters = hitters.rename(columns={"yearID": "Season"})
    pit = pit.rename(columns={"yearID": "Season"})

    hitter_columns = ["Season", "section", "rank", "Player", "qualifying_stat",
                      "Position", "Common stats", "G", "AB", "H", "HR", "RBI", "AVG", "OPS",
                      "doubles", "triples", "BB", "HBP", "SF", "playerID"]
    pitcher_columns = ["Season", "section", "rank", "Player", "qualifying_stat",
                       "Common stats", "G", "GS", "IP", "IPOuts", "W", "L", "ER", "ERA", "SO", "SV", "playerID"]
    hitters_out = hitters[hitter_columns].copy()
    pit_out = pit[pitcher_columns].copy()
    all_rows = pd.concat([hitters_out, pit_out], ignore_index=True, sort=False)
    all_rows = all_rows.sort_values(["Season", "section", "rank"]).reset_index(drop=True)

    # The physical binder position is the pocket on that season's 3-page set.
    def slot_for(row):
        rank = int(row["rank"])
        if row["section"] == "Pitcher":
            page = 3
            pocket_rank = rank
        elif rank <= 9:
            page = 1
            pocket_rank = rank
        else:
            page = 2
            pocket_rank = rank - 9
        row_num = (pocket_rank - 1) // 3
        col_num = (pocket_rank - 1) % 3 + 1
        row_letter = "ABC"[row_num]
        return f'{int(row["Season"])}-{page}-{row_letter}{col_num}'

    all_rows["Binder page / slot"] = [
        slot_for(row) for _, row in all_rows.iterrows()
    ]
    all_rows["Preferred card year"] = all_rows["Season"] + 1
    all_rows["Card owned?"] = ""
    all_rows["Actual card year"] = ""
    all_rows["Brand / set"] = ""
    all_rows["Card number"] = ""
    all_rows["Team uniform? (Y/N/Unsure)"] = ""
    all_rows["Shows prior-season stats? (Y/N/Unsure)"] = ""
    all_rows["Cost"] = ""
    all_rows["Order ID"] = ""
    all_rows["Notes / preference exception"] = ""
    return all_rows, teams, managers, people


def aggregate_decade_rows(data, years):
    """Aggregate the displayed season leaders into one alternate season sheet."""
    decade = data[data["Season"].isin(years)].copy()
    aggregate_rows = []

    hitters = decade[decade["section"] == "Position player"]
    if not hitters.empty:
        grouped = hitters.groupby("playerID", dropna=False)
        for player_id, group in grouped:
            totals = {column: group[column].fillna(0).sum() for column in
                      ["G", "AB", "H", "HR", "RBI", "doubles", "triples", "BB", "HBP", "SF"]}
            at_bats = totals["AB"]
            plate_appearances = at_bats + totals["BB"] + totals["HBP"] + totals["SF"]
            on_base = totals["H"] + totals["BB"] + totals["HBP"]
            slugging = (totals["H"] + 2 * totals["doubles"] + 3 * totals["triples"] +
                        3 * totals["HR"]) / at_bats if at_bats else 0
            aggregate_rows.append({
                "section": "Position player", "Player": group.iloc[0]["Player"],
                "playerID": player_id, "G": totals["G"], "AB": at_bats,
                "H": totals["H"], "HR": totals["HR"], "RBI": totals["RBI"],
                "AVG": totals["H"] / at_bats if at_bats else 0,
                "OPS": on_base / plate_appearances + slugging if plate_appearances else 0,
                "Position": ", ".join(dict.fromkeys(
                    position.strip()
                    for positions in group["Position"]
                    for position in str(positions).split(",")
                    if position.strip()
                )),
            })
        aggregate_rows[-len(hitters["playerID"].unique()):] = sorted(
            aggregate_rows[-len(hitters["playerID"].unique()):],
            key=lambda row: (-row["G"], -row["AB"], str(row["playerID"])),
        )
        for rank, row in enumerate(aggregate_rows[:18], start=1):
            row["rank"] = rank
        aggregate_rows = aggregate_rows[:18]

    pitchers = decade[decade["section"] == "Pitcher"]
    pitcher_rows = []
    if not pitchers.empty:
        for player_id, group in pitchers.groupby("playerID", dropna=False):
            totals = {column: group[column].fillna(0).sum() for column in
                      ["G", "GS", "IPOuts", "W", "L", "ER", "SO", "SV"]}
            pitcher_rows.append({
                "section": "Pitcher", "Player": group.iloc[0]["Player"],
                "playerID": player_id, **totals,
                "IP": totals["IPOuts"] / 3,
                "ERA": totals["ER"] * 27 / totals["IPOuts"] if totals["IPOuts"] else 0,
            })
        pitcher_rows.sort(key=lambda row: (-row["IPOuts"], -row["SO"], str(row["playerID"])))
        for rank, row in enumerate(pitcher_rows[:9], start=1):
            row["rank"] = rank
        aggregate_rows.extend(pitcher_rows[:9])

    return pd.DataFrame(aggregate_rows)

def format_innings(outs):
    outs = int(outs)
    full = outs // 3
    remainder = outs % 3
    return f"{full}.{remainder} IP"


def style_table(table, header=True, font_size=8, padding=5):
    commands = [
        ("FONTNAME", (0, 0), (-1, -1), PDF_FONT),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#B8C0CC")),
        ("LEFTPADDING", (0, 0), (-1, -1), padding),
        ("RIGHTPADDING", (0, 0), (-1, -1), padding),
        ("TOPPADDING", (0, 0), (-1, -1), padding),
        ("BOTTOMPADDING", (0, 0), (-1, -1), padding),
    ]
    if header:
        commands += [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D5DCE5")),
            ("FONTNAME", (0, 0), (-1, 0), PDF_FONT_BOLD),
        ]
    for row in range(2, len(table._cellvalues), 2):
        commands.append(("BACKGROUND", (0, row), (-1, row), colors.HexColor("#F5F7FA")))
    table.setStyle(TableStyle(commands))
    return table


def empty_checkbox():
    drawing = Drawing(7, 7)
    drawing.add(Rect(0.5, 0.5, 6, 6, fillColor=None,
                    strokeColor=colors.HexColor("#333333"), strokeWidth=0.8))
    return drawing


def safe_filename(team_name):
    filename = re.sub(r"[^a-z0-9]+", "-", team_name.lower()).strip("-")
    return f"{filename or 'team'}.pdf"


def create_pdf(
    data, path, start, end, team_ids, team_name, team_seasons,
    teams, managers, people
):
    global PDF_FONT, PDF_FONT_BOLD
    font_dir = Path.home() / "AppData" / "Local" / "Microsoft" / "Windows" / "Fonts"
    roboto_regular = font_dir / "Roboto-Regular.ttf"
    roboto_bold = font_dir / "Roboto-Bold.ttf"
    if roboto_regular.exists() and roboto_bold.exists():
        try:
            pdfmetrics.registerFont(TTFont("Roboto", str(roboto_regular)))
            pdfmetrics.registerFont(TTFont("Roboto-Bold", str(roboto_bold)))
            registerFontFamily(
                "Roboto", normal="Roboto", bold="Roboto-Bold",
                italic="Roboto", boldItalic="Roboto-Bold"
            )
            PDF_FONT, PDF_FONT_BOLD = "Roboto", "Roboto-Bold"
        except Exception:
            PDF_FONT, PDF_FONT_BOLD = "Helvetica", "Helvetica-Bold"
    styles = getSampleStyleSheet()
    styles["Normal"].fontName = PDF_FONT
    styles["Title"].fontName = PDF_FONT_BOLD
    styles["Title"].fontSize = 21
    styles["Title"].leading = 19
    styles["Title"].alignment = 0
    styles.add(ParagraphStyle(
        name="TeamTitle", parent=styles["Title"], fontSize=11, leading=12,
        spaceBefore=0, spaceAfter=0
    ))
    styles["Heading2"].fontName = PDF_FONT_BOLD
    styles["Heading2"].fontSize = 11
    styles["Heading2"].leading = 13
    doc = SimpleDocTemplate(
        str(path), pagesize=letter,
        rightMargin=0.3 * inch, leftMargin=0.3 * inch,
        topMargin=0.25 * inch, bottomMargin=0.25 * inch
    )
    story = []
    people_names = people.copy()
    people_names["Manager"] = (
        people_names["nameFirst"].fillna("").astype(str).str.strip() + " " +
        people_names["nameLast"].fillna("").astype(str).str.strip()
    ).str.strip()
    manager_name_map = dict(zip(people_names["playerID"], people_names["Manager"]))

    def value(row, column, default=""):
        result = row.get(column, default)
        return default if pd.isna(result) else result

    included_seasons = [season for season in range(end, start - 1, -1) if season in team_seasons]
    seasons = [season for season in included_seasons if not 1960 <= season <= 1969]
    for season_index, season in enumerate(seasons):
        season_rows = data[data["Season"] == season]
        team_rows = teams[
            (teams["yearID"] == season) &
            (teams["teamID"].isin(team_ids))
        ]
        team_info = team_rows.iloc[0] if not team_rows.empty else None
        if team_info is not None:
            wins = int(value(team_info, "W", 0))
            losses = int(value(team_info, "L", 0))
            total_games = wins + losses
            winning_percentage = wins / total_games if total_games else 0
            winning_percentage_text = f"{winning_percentage:.3f}".lstrip("0")
            record = f"{wins}-{losses}"
            stadium_games = {}
            for _, team_row in team_rows.iterrows():
                park = str(value(team_row, "park", "Not listed in database"))
                games_at_park = value(team_row, "G", "")
                if pd.notna(games_at_park) and str(games_at_park).strip():
                    stadium_games[park] = stadium_games.get(park, 0) + int(games_at_park)
                else:
                    stadium_games.setdefault(park, 0)
            if len(stadium_games) > 1:
                home_field = "<br/>".join(
                    f"<b>{park}</b> ({game_count} G)"
                    for park, game_count in stadium_games.items()
                )
            else:
                park = next(iter(stadium_games), "Not listed in database")
                home_field = f"<b>{park}</b>"
            division_rank = value(team_info, "rank")
            games = value(team_info, "G")
            attendance = value(team_info, "attendance")
        else:
            record, winning_percentage_text, home_field = "Not available", "", "Not listed"
            division_rank, games, attendance = "", "", ""

        mgr_rows = managers[
            (managers["yearID"] == season) &
            (managers["teamID"].isin(team_ids))
        ].copy()
        manager_labels = []
        if not mgr_rows.empty:
            mgr_rows = mgr_rows.sort_values([c for c in ["inseason"] if c in mgr_rows.columns])
            for _, manager in mgr_rows.iterrows():
                manager_id = manager.get("playerID", "")
                name = manager_name_map.get(manager_id, str(manager_id))
                games_managed = manager.get("G", "")
                if pd.notna(games_managed) and str(games_managed).strip():
                    managed_games = int(games_managed)
                    team_games = int(games) if games != "" and pd.notna(games) else None
                    if managed_games > 1 and managed_games != team_games:
                        manager_labels.append(f"<b>{name}</b> ({managed_games} G)")
                    else:
                        manager_labels.append(f"<b>{name}</b>")
                else:
                    manager_labels.append(f"<b>{name}</b>")
        manager_text = "<br/>".join(manager_labels) if manager_labels else "Not listed in database"

        story.append(Paragraph(f"{season} Season", styles["Title"]))
        story.append(Spacer(1, 3))
        story.append(Paragraph(team_name.upper(), styles["TeamTitle"]))
        story.append(Spacer(1, 10))
        story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor("#333333"), spaceBefore=0, spaceAfter=20))
        stat_items = [
            ("RECORD", record),
            ("MANAGER", manager_text),
            ("HOME FIELD", home_field),
        ]
        if division_rank != "" and pd.notna(division_rank):
            stat_items.append(("DIVISION FINISH", division_rank))
        if games != "" and pd.notna(games):
            stat_items.append(("TEAM GAMES", int(games)))
        if attendance != "" and pd.notna(attendance):
            try:
                stat_items.append(("ATTENDANCE", f"{int(attendance):,}"))
            except (ValueError, TypeError):
                pass
        stat_cells = []
        fixed_widths = {"MANAGER": 1.8, "HOME FIELD": 1.65}
        remaining_width = 7.5 - sum(fixed_widths.get(label, 0) for label, _ in stat_items)
        flexible_count = sum(label not in fixed_widths for label, _ in stat_items)
        stat_widths = [
            fixed_widths.get(label, remaining_width / flexible_count)
            for label, _ in stat_items
        ]
        for label, stat_value in stat_items:
            if label == "RECORD":
                value_markup = (
                    f"<b>{stat_value}</b> ({winning_percentage_text})"
                    if winning_percentage_text else f"<b>{stat_value}</b>"
                )
            elif label in {"MANAGER", "HOME FIELD"}:
                value_markup = stat_value
            else:
                value_markup = f"<b>{stat_value}</b>"
            stat_cells.append(Table([
                [Paragraph(f'<font color="#8A929D" size="7">{label}</font>', styles["Normal"])],
                [Paragraph(f'<font color="#111111" size="10">{value_markup}</font>', styles["Normal"])],
            ], colWidths=[stat_widths[len(stat_cells)] * inch], style=TableStyle([
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, 0), 0),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 1),
                ("TOPPADDING", (0, 1), (-1, 1), 0),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 0),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ])))
        eyebrow_table = Table([stat_cells], colWidths=[width * inch for width in stat_widths], hAlign="LEFT", style=TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ]))
        story.append(eyebrow_table)
        story.append(Spacer(1, 8))

        table_specs = [
            ("POSITION PLAYERS", "Position player", range(1, 19),
             ["", "SLOT", "PLAYER", "POSITION", "G", "AB", "H", "HR", "RBI", "AVG", "OPS"],
             ["G", "AB", "H", "HR", "RBI", "AVG", "OPS"]),
            ("PITCHERS", "Pitcher", range(1, 10),
             ["", "SLOT", "PLAYER", "IP", "G", "GS", "W", "L", "ERA", "SO", "SV"],
             ["IP", "G", "GS", "W", "L", "ERA", "SO", "SV"]),
        ]
        for title, section, ranks, headers, stat_columns in table_specs:
            story.append(Paragraph(title, styles["Heading2"]))
            subset = season_rows[season_rows["section"] == section]
            by_rank = {int(row["rank"]): row for _, row in subset.iterrows()}
            table_data = [headers]
            for rank in ranks:
                row = by_rank.get(rank)
                pocket_rank = rank if section == "Pitcher" else (rank if rank <= 9 else rank - 9)
                page_num = 3 if section == "Pitcher" else (1 if rank <= 9 else 2)
                row_letter = "ABC"[(pocket_rank - 1) // 3]
                col_num = (pocket_rank - 1) % 3 + 1
                slot = f"{page_num}-{row_letter}{col_num}"
                if row is None:
                    values = [empty_checkbox(), slot, "________________", "—"] + ["—"] * len(stat_columns)
                else:
                    values = [empty_checkbox(), slot, str(row["Player"])]
                    if section == "Position player":
                        values.append(str(row.get("Position", "")))
                    for column in stat_columns:
                        item = value(row, column, "")
                        if column in {"G", "AB", "H", "HR", "RBI", "GS", "W", "L", "ER", "SO", "SV"} and item != "":
                            item = str(int(item))
                        elif column == "IP" and item != "":
                            item = format_innings(int(round(float(item) * 3))).replace(" IP", "")
                        elif column in {"AVG", "OPS", "ERA"} and item != "":
                            decimals = 2 if column == "ERA" else 3
                            item = f"{float(item):.{decimals}f}"
                            if column == "AVG":
                                item = item.replace("0.", ".")
                        values.append(str(item))
                table_data.append(values)
            base_widths = ([0.3, 0.43, 1.55] +
                           ([0.9] if section == "Position player" else []) +
                           [0.38] * len(stat_columns))
            widths = [base_widths[0]] + [width + 0.08 for width in base_widths[1:]]
            tbl = Table(table_data, colWidths=[width * inch for width in widths],
                        rowHeights=[0.25 * inch] + [0.22 * inch] * len(list(ranks)),
                        repeatRows=1, hAlign="LEFT")
            styled_table = style_table(tbl, font_size=8.2, padding=5)
            styled_table.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 1), (0, -1), "CENTER"),
                ("RIGHTPADDING", (0, 0), (0, -1), 5),
                ("RIGHTPADDING", (1, 0), (-1, -1), 8),
            ]))
            story.append(styled_table)
            story.append(Spacer(1, 10))
        if season_index < len(seasons) - 1:
            story.append(PageBreak())

    decade_specs = [
        ("1970-79 Season", range(1970, 1980)),
        ("1960-69 Season", range(1960, 1970)),
    ]
    for decade_title, decade_range in decade_specs:
        decade_years = [season for season in included_seasons if season in decade_range]
        decade_data = aggregate_decade_rows(data, decade_years)
        if decade_data.empty:
            continue
        story.append(PageBreak())
        story.append(Paragraph(decade_title, styles["Title"]))
        story.append(Spacer(1, 3))
        story.append(Paragraph(team_name.upper(), styles["TeamTitle"]))
        story.append(Spacer(1, 10))
        story.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor("#333333"), spaceBefore=0, spaceAfter=20))

        table_specs = [
            ("POSITION PLAYERS", "Position player", range(1, 19),
             ["", "SLOT", "PLAYER", "POSITION", "G", "AB", "H", "HR", "RBI", "AVG", "OPS"],
             ["G", "AB", "H", "HR", "RBI", "AVG", "OPS"]),
            ("PITCHERS", "Pitcher", range(1, 10),
             ["", "SLOT", "PLAYER", "IP", "G", "GS", "W", "L", "ERA", "SO", "SV"],
             ["IP", "G", "GS", "W", "L", "ERA", "SO", "SV"]),
        ]
        for title, section, ranks, headers, stat_columns in table_specs:
            story.append(Paragraph(title, styles["Heading2"]))
            subset = decade_data[decade_data["section"] == section]
            by_rank = {int(row["rank"]): row for _, row in subset.iterrows()}
            table_data = [headers]
            for rank in ranks:
                row = by_rank.get(rank)
                pocket_rank = rank if section == "Pitcher" else (rank if rank <= 9 else rank - 9)
                page_num = 3 if section == "Pitcher" else (1 if rank <= 9 else 2)
                row_letter = "ABC"[(pocket_rank - 1) // 3]
                col_num = (pocket_rank - 1) % 3 + 1
                slot = f"{page_num}-{row_letter}{col_num}"
                if row is None:
                    values = [empty_checkbox(), slot, "________________", "—"] + ["—"] * len(stat_columns)
                else:
                    values = [empty_checkbox(), slot, str(row["Player"])]
                    if section == "Position player":
                        values.append(str(row.get("Position", "")))
                    for column in stat_columns:
                        item = value(row, column, "")
                        if column in {"G", "AB", "H", "HR", "RBI", "GS", "W", "L", "SO", "SV"} and item != "":
                            item = str(int(item))
                        elif column == "IP" and item != "":
                            item = format_innings(int(round(float(row["IPOuts"])))) .replace(" IP", "")
                        elif column in {"AVG", "OPS", "ERA"} and item != "":
                            decimals = 2 if column == "ERA" else 3
                            item = f"{float(item):.{decimals}f}"
                            if column == "AVG":
                                item = item.replace("0.", ".")
                        values.append(str(item))
                table_data.append(values)
            base_widths = ([0.3, 0.43, 1.55] +
                           ([0.9] if section == "Position player" else []) +
                           [0.38] * len(stat_columns))
            widths = [base_widths[0]] + [width + 0.08 for width in base_widths[1:]]
            tbl = Table(table_data, colWidths=[width * inch for width in widths],
                        rowHeights=[0.25 * inch] + [0.22 * inch] * len(list(ranks)),
                        repeatRows=1, hAlign="LEFT")
            styled_table = style_table(tbl, font_size=8.2, padding=5)
            styled_table.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 1), (0, -1), "CENTER"),
                ("RIGHTPADDING", (0, 0), (0, -1), 5),
                ("RIGHTPADDING", (1, 0), (-1, -1), 8),
            ]))
            story.append(styled_table)
            story.append(Spacer(1, 10))
    doc.build(story)

def main():
    parser = argparse.ArgumentParser(
        description="Build season card checklists for MLB teams (AL and NL)."
    )
    parser.add_argument("--team-id", default=None,
                        help="Only generate this MLB Lahman teamID; default is all MLB teams")
    parser.add_argument("--start", type=int, default=1960,
                        help="Earliest season to include (default: 1960)")
    parser.add_argument("--end", type=int, default=None,
                        help="Latest season to include; default is each team's latest available season")
    parser.add_argument("--output", default="checklists",
                        help="Output directory (default: ./checklists)")
    args = parser.parse_args()

    print("Loading Lahman baseball data...")
    batting, pitching, people, appearances, teams, managers = get_lahman_tables()
    outdir = Path(args.output)
    outdir.mkdir(parents=True, exist_ok=True)

    # CSV/XLSX export is intentionally disabled; the printable PDF is the only output.
    # csv_path = outdir / "team_card_checklists.csv"
    # data.to_csv(csv_path, index=False)
    # xlsx_path = outdir / "team_card_checklists.xlsx"
    # with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    #     data.to_excel(writer, sheet_name="Master Checklist", index=False)

    teams = teams.copy()
    teams["yearID"] = pd.to_numeric(teams["yearID"], errors="coerce")
    team_years = teams[
        teams["lgID"].isin(MLB_LEAGUES)
    ].dropna(subset=["yearID"])
    team_groups = []
    for franchise_id, group in team_years.groupby("franchID", dropna=False):
        group = group.sort_values("yearID")
        team_ids = set(group["teamID"].astype(str))
        team_name = str(group.iloc[-1]["name"])
        team_groups.append((str(franchise_id), team_ids, team_name, group))
    if args.team_id:
        team_id = args.team_id.upper()
        selected_groups = [
            group for group in team_groups
            if group[0] == team_id or team_id in group[1]
        ]
        unknown_ids = [] if selected_groups else [team_id]
    else:
        selected_groups = sorted(team_groups, key=lambda item: (item[2], item[0]))
        unknown_ids = []
    if unknown_ids:
        raise SystemExit(f"Unknown Lahman teamID(s): {', '.join(unknown_ids)}")

    used_filenames = set()
    generated = []
    for franchise_id, team_ids, team_name, group in selected_groups:
        team_season_values = sorted(int(year) for year in group["yearID"].unique())
        first_season = team_season_values[0]
        last_season = team_season_values[-1]
        start = max(first_season, args.start)
        end = min(last_season, args.end) if args.end is not None else last_season
        if start > end:
            continue

        team_seasons = {year for year in team_season_values if start <= year <= end}
        filename = safe_filename(team_name)
        if filename in used_filenames:
            filename = safe_filename(f"{team_name}-{franchise_id}")
        used_filenames.add(filename)
        pdf_path = outdir / filename

        data, _, _, _ = prepare_data(
            batting, pitching, people, appearances, teams, managers,
            team_ids, start, end, team_seasons
        )
        create_pdf(
            data, pdf_path, start, end, team_ids, team_name, team_seasons,
            teams, managers, people
        )
        generated.append((team_name, start, end, pdf_path))

    print("\nDone. Files created:")
    for team_name, start, end, pdf_path in generated:
        print(f"  PDF:   {pdf_path} ({team_name}, {start}-{end})")



if __name__ == "__main__":
    main()
