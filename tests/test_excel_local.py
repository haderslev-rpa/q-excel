from pathlib import Path
from q_excel.edit_excel import excel_edit


def test_read_ark2():
    print("Starter test: Læs fra Ark2")

    file_path = Path("exceltest.xlsx")

    value = excel_edit(
        file_path=str(file_path),
        mode="read",
        from_ref="A1",
        sheet_name="Ark2"
    )

    print(f"Værdi i Ark2!A1: {value}")

    assert value is not None, "A1 i Ark2 er tom eller findes ikke"

    print("✅ TEST BESTÅET")


if __name__ == "__main__":
    test_read_ark2()