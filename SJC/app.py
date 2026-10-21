from pathlib import Path
import os

from dash import Dash, Input, Output, ctx, dcc, html


REVIEW_FOLDER = Path(
    os.environ.get("SJC_REVIEW_FOLDER", Path.home() / "SJA" / "Review")
)
CHARTS = [
    ("nuts", "Nuts"),
    ("dgw1", "DGW1"),
    ("gate", "Gate"),
    ("speed", "Speed"),
    ("wind", "Wind"),
    ("distance", "Distance"),
    ("dgw2", "DGW2"),
    ("judge", "Judge"),
]
RANKS = range(1, 10)


def find_workbooks() -> list[Path]:
    if not REVIEW_FOLDER.is_dir():
        return []
    return sorted(
        (path for path in REVIEW_FOLDER.glob("*.xlsx") if not path.name.startswith("~$")),
        key=lambda path: path.name.casefold(),
        reverse=True,
    )


workbooks = find_workbooks()
workbook_options = [
    {"label": workbook.stem, "value": str(workbook)} for workbook in workbooks
]
initial_workbook = workbook_options[0]["value"] if workbook_options else None
chart_inputs = [Input(f"chart-{name}", "n_clicks") for name, _ in CHARTS]
rank_inputs = [Input(f"rank-{rank}", "n_clicks") for rank in RANKS]
all_button_inputs = chart_inputs + rank_inputs
all_button_ids = [f"chart-{name}" for name, _ in CHARTS] + [
    f"rank-{rank}" for rank in RANKS
]

app = Dash(__name__, title="SJC – Ski Jumping Charts")
server = app.server

app.layout = html.Div(
    className="app-shell",
    children=[
        dcc.Store(id="selected-chart", data={"name": "nuts", "rank": None}),
        html.Header(
            className="topbar",
            children=[
                html.Div("SJC – Ski Jumping Charts", className="brand"),
                html.Div(
                    id="competition-title",
                    className="competition-title",
                    children="Wettkampf auswählen" if not initial_workbook else "",
                ),
                html.Div(
                    className="competition-picker",
                    children=[
                        dcc.Dropdown(
                            id="workbook-picker",
                            options=workbook_options,
                            value=initial_workbook,
                            placeholder="Wettkampf auswählen",
                            clearable=False,
                            searchable=True,
                            className="workbook-dropdown",
                        ),
                        html.Div(
                            "Keine .xlsx-Dateien im Review-Ordner gefunden."
                            if not workbooks
                            else f"{len(workbooks)} Wettkampf"
                            f"{'' if len(workbooks) == 1 else 'e'} verfügbar",
                            className="picker-hint",
                        ),
                    ],
                ),
            ],
        ),
        html.Div(
            className="content-shell",
            children=[
                html.Aside(
                    className="sidebar",
                    children=[
                        html.Nav(
                            className="chart-navigation",
                            children=[
                                html.Button(
                                    label,
                                    id=f"chart-{name}",
                                    n_clicks=0,
                                    className=(
                                        "chart-button active"
                                        if name == "nuts"
                                        else "chart-button"
                                    ),
                                )
                                for name, label in CHARTS[:-1]
                            ],
                        ),
                        html.Button(
                            "Judge",
                            id="chart-judge",
                            n_clicks=0,
                            className="chart-button judge-button",
                        ),
                        html.Div(
                            className="rank-grid",
                            children=[
                                html.Button(
                                    str(rank),
                                    id=f"rank-{rank}",
                                    n_clicks=0,
                                    className="rank-button",
                                    title=f"Judge Rangliste {rank}",
                                    **{"aria-label": f"Judge Rangliste {rank}"},
                                )
                                for rank in RANKS
                            ],
                        ),
                    ],
                ),
                html.Main(
                    id="workspace",
                    className="workspace",
                ),
            ],
        ),
    ],
)


@app.callback(
    Output("competition-title", "children"),
    Input("workbook-picker", "value"),
)
def show_competition_title(workbook_path: str | None) -> str:
    return Path(workbook_path).stem if workbook_path else "Wettkampf auswählen"


@app.callback(
    Output("selected-chart", "data"),
    all_button_inputs,
    prevent_initial_call=True,
)
def select_chart(*_clicks: int | None) -> dict[str, str | int | None]:
    button_id = ctx.triggered_id
    if not isinstance(button_id, str):
        return {"name": "nuts", "rank": None}
    if button_id.startswith("rank-"):
        return {"name": "judge", "rank": int(button_id.removeprefix("rank-"))}
    return {"name": button_id.removeprefix("chart-"), "rank": None}


@app.callback(
    [Output(button_id, "className") for button_id in all_button_ids]
    + [Output("workspace", "children")],
    Input("selected-chart", "data"),
    Input("workbook-picker", "value"),
)
def render_selection(
    selection: dict[str, str | int | None], workbook_path: str | None
) -> list[object]:
    selected_name = selection.get("name", "nuts")
    selected_rank = selection.get("rank")
    classes = [
        (
            "chart-button active"
            if button_id == f"chart-{selected_name}" and selected_rank is None
            else "chart-button"
        )
        if button_id.startswith("chart-")
        else (
            "rank-button active"
            if selected_name == "judge"
            and selected_rank == int(button_id.removeprefix("rank-"))
            else "rank-button"
        )
        for button_id in all_button_ids
    ]
    chart_label = dict(CHARTS).get(str(selected_name), "Judge")
    heading = (
        f"Judge – Rangliste {selected_rank}"
        if selected_name == "judge" and selected_rank is not None
        else chart_label
    )
    workbook_name = Path(workbook_path).stem if workbook_path else None
    workspace = html.Div(
        className="workspace-empty",
        children=[
            html.Div(className="workspace-mark", children="SJC"),
            html.Div(heading, className="workspace-title"),
            html.Div(
                workbook_name or "Wähle rechts oben eine Wettkampfdatei aus.",
                className="workspace-caption",
            ),
            html.Div(
                "Hier erscheinen später die interaktiven Ranglisten-Grafiken.",
                className="workspace-description",
            ),
        ],
    )
    return classes + [workspace]


if __name__ == "__main__":
    app.run(debug=True)
