from typing import List
import jinja2
from dataclasses import dataclass
import json
from pathlib import Path
import xlsxwriter


RESULTS_PATH = Path("RESULTS/results.json")
RESULTS_OUT_MD_PATH = Path("RESULTS/README.md")
RESULTS_OUT_XLSX_PATH = "RESULTS/results.xlsx"


@dataclass
class ResultLine:
    vendor: str
    application_id: str
    version: str
    doc_type: List[str]
    service: List[str]
    date: str
    gtw_version: str

    def md_table_line(self) -> str:
        return (
            "|"
            + "|".join(
                [
                    self.vendor.replace("|", r"\|"),
                    self.application_id.replace("|", r"\|"),
                    self.version.replace("|", r"\|"),
                    ",".join(map(lambda x: x.replace("|", r"\|"), self.doc_type)),
                    ",".join(map(lambda x: x.replace("|", r"\|"), self.service)),
                    self.date.replace("|", r"\|"),
                    self.gtw_version.replace("|", r"\|"),
                ]
            )
            + "|"
        )

    def flatten_line(self):
        return [
            self.vendor,
            self.application_id,
            self.version,
            ",".join(self.doc_type),
            ",".join(self.service),
            self.date,
            self.gtw_version,
        ]


def generate_md(md_table_lines: List[str]):
    print(md_lines)
    templateLoader = jinja2.FileSystemLoader(searchpath="./")
    templateEnv = jinja2.Environment(loader=templateLoader)
    TEMPLATE_FILE = "RESULTS.md.tpl"
    template = templateEnv.get_template(TEMPLATE_FILE)
    outputText = template.render(md_table_lines=md_table_lines)
    RESULTS_OUT_MD_PATH.write_text(outputText, encoding="utf8")


def generate_xlsx(xls_lines: List[List[str]]):
    workbook = xlsxwriter.Workbook(RESULTS_OUT_XLSX_PATH)
    worksheet = workbook.add_worksheet()

    for row_num, row_data in enumerate(xls_lines):
        for col_num, col_data in enumerate(row_data):
            worksheet.write(row_num, col_num, col_data)

    for col in range(len(xls_lines[0])):
        worksheet.set_column(col,col,width=max([len(xls_lines[r][col]) for r in range(len(xls_lines))]))

    workbook.close()


if __name__ == "__main__":
    try:
        with RESULTS_PATH.open("r", encoding="utf8") as results_file:
            data = json.load(results_file)
        md_lines = []
        xls_lines = []
        xls_lines.append(
            [
                "Fornitore",
                "Applicativo",
                "Versione",
                "Tipo Documento",
                "Servizio",
                "Data validazione",
                "Versione Gateway",
            ]
        )
        for d in data["results"]:
            rl = ResultLine(**d)
            md_lines.append(rl.md_table_line())
            xls_lines.append(rl.flatten_line())
        generate_md(md_lines)
        generate_xlsx(xls_lines)
    except:
        raise
