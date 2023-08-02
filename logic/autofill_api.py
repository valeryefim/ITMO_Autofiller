from datetime import datetime
from docx import Document
import requests
from urllib.parse import urlparse


class UserDocument:
    def __init__(self, doc_type, number, given_by, issue_date):
        self.type = doc_type
        self.number = number
        self.given_by = given_by
        self.issue_date = issue_date


class User:
    def __init__(
        self,
        name,
        email,
        phone,
        birth_date,
        doc_type,
        number,
        given_by,
        issue_date,
        address,
        degree,
        program,
        direction,
    ):
        self.name = name
        self.email = email
        self.phone = phone
        self.birth_date = birth_date
        self.document = UserDocument(doc_type, number, given_by, issue_date)
        self.address = address
        self.degree = degree
        self.program = program
        self.direction = direction


class Customer:
    def __init__(
        self,
        payment_type,
        name=None,
        address=None,
        doc_type=None,
        number=None,
        given_by=None,
        phone=None,
        document=None,
        document_date=None,
        ogrn=None,
        inn=None,
        behalf=None,
        behalf_genitive=None,
    ):
        self.payment_type = payment_type
        if payment_type == "individual":
            self.name = name
            self.document = UserDocument(doc_type, number, given_by, issue_date=None)
            self.phone = phone
            self.address = address
        elif payment_type == "juridical":
            self.address = address
            self.name = name
            self.document = (document,)
            self.document_date = document_date
            self.ogrn = ogrn
            self.inn = inn
            self.behalf = behalf
            self.behalf_genitive = behalf_genitive


def get_program_by_financing(selected_programs):
    idx = 0
    for program in selected_programs:
        if program["financing"] == "contract":
            return idx, program["program"]
        idx += 1
    return None, None


def get_data(cookies, current_url):
    parsed_url = urlparse(current_url)
    path_parts = parsed_url.path.split("/")
    user_id = path_parts[4]
    details_id = path_parts[5]

    application = requests.get(
        f"https://abitlk.itmo.ru/api/v1/users/{user_id}/studentDetails/" f"{details_id}/application",
        headers={"Cookie": cookies},
    ).json()
    personally = requests.get(
        f"https://abitlk.itmo.ru/api/v1/users/{user_id}/studentDetails/" f"{details_id}/forms/personally",
        headers={"Cookie": cookies},
    ).json()
    address = requests.get(
        f"https://abitlk.itmo.ru/api/v1/users/{user_id}/studentDetails/" f"{details_id}/forms/address",
        headers={"Cookie": cookies},
    ).json()
    programs = requests.get(
        f"https://abitlk.itmo.ru/api/v1/users/{user_id}/studentDetails/" f"{details_id}/programs",
        headers={"Cookie": cookies},
    ).json()
    payment = requests.get(
        f"https://abitlk.itmo.ru/api/v1/users/{user_id}/studentDetails/" f"{details_id}/forms/payment",
        headers={"Cookie": cookies},
    ).json()

    for response in (application, personally, address, programs, payment):
        if not response["ok"]:
            raise ConnectionError(response["message"])

    contract_idx, contract_program = get_program_by_financing(application["result"]["selected_programs"])

    user = User(
        name=application["result"]["user"]["full_name"],
        email=application["result"]["user"]["email"],
        phone=application["result"]["user"]["phone"],
        birth_date=application["result"]["user"]["birth_date"],
        doc_type=personally["result"]["documents"]["person_id"]["type"],
        number=application["result"]["user"]["passport_number"],
        given_by=personally["result"]["documents"]["person_id"]["division_name"],
        issue_date=personally["result"]["documents"]["person_id"]["issued_date"],
        address=address["result"]["registration_address"],
        degree=application["result"]["user"]["degree"],
        program=contract_program,
        direction=programs["result"]["selected_programs"][contract_idx]["competitive_group"]["title"]
        if application["result"]["user"]["degree"] == "bachelor"
        else programs["result"]["selected_programs"][contract_idx]["program"]["direction_of_education"],
    )
    user.birth_date = datetime.fromisoformat(user.birth_date.replace("Z", "+00:00"))
    user.document.issue_date = datetime.fromisoformat(user.document.issue_date.replace("Z", "+00:00"))

    if payment["result"]["payment_type"] == "individual":
        customer = Customer(
            payment_type=payment["result"]["payment_type"],
            name=payment["result"]["individual"]["full_name"]["first_name"]
            + " "
            + payment["result"]["individual"]["full_name"]["last_name"]
            + " "
            + payment["result"]["individual"]["full_name"]["middle_name"],
            address=payment["result"]["individual"]["address"],
            doc_type="Паспорт гражданина РФ",
            number=payment["result"]["individual"]["series"] + payment["result"]["individual"]["number"],
            given_by=payment["result"]["individual"]["division_name"],
            phone=payment["result"]["individual"]["phone"],
        )
    elif payment["result"]["payment_type"] == "juridical":
        customer = Customer(
            payment_type=payment["result"]["payment_type"],
            name=payment["result"]["juridical"]["name"],
            address=payment["result"]["juridical"]["address"],
            document=payment["result"]["juridical"]["document_name_genitive"],
            document_date=payment["result"]["juridical"]["date_document"],
            ogrn=payment["result"]["juridical"]["ogrn"],
            inn=payment["result"]["juridical"]["inn"],
            behalf=payment["result"]["juridical"]["full_name_behalf"],
            behalf_genitive=payment["result"]["juridical"]["full_name_behalf_genitive"],
        )
        customer.document_date = datetime.fromisoformat(customer.document_date.replace("Z", "+00:00"))
    else:
        customer = Customer(payment_type=payment["result"]["payment_type"])

    return user, customer


def fill_contract(user, customer, contract_path="../logic/data/sample_contract_api.docx"):
    annual = {
        "Бизнес-информатика": "301000 (Триста одна тысяча)",
        "Технологии и инновации": "301000 (Триста одна тысяча)",
        "Управление высокотехнологичным бизнесом": "349000 (Триста сорок девять тысяч)",
        "Технологии и стратегии бизнес-трансформации": "349000 (Триста сорок девять тысяч)",
        "Цифровые продукты: создание и управление": "380000 (Триста восемьдесят тысяч)",
        "Стратегическое управление интеллектуальной собственностью / IP Management Strategy": "349000 (Триста сорок "
        "девять тысяч)",
    }
    total = {
        "Бизнес-информатика": "1204000 (Один миллион двести четыре тысячи)",
        "Технологии и инновации": "1204000 (Один миллион двести четыре тысячи)",
        "Управление высокотехнологичным бизнесом": "698000 (Шестьсот девяносто восемь)",
        "Технологии и стратегии бизнес-трансформации": "698000 (Шестьсот девяносто восемь)",
        "Цифровые продукты: создание и управление": "760000 (Семьсот шестьдесят тысяч)",
        "Стратегическое управление интеллектуальной собственностью / IP Management Strategy": "698000 (Шестьсот "
        "девяносто восемь)",
    }

    contract = Document(contract_path)

    today = datetime.today()
    user_age = (
        today.year - user.birth_date.year - ((today.month, today.day) < (user.birth_date.month, user.birth_date.day))
    )
    end_flag = False

    for paragraph in contract.paragraphs:
        for run in paragraph.runs:
            if "DATE" in run.text:
                months = {7: "июля", 8: "августа"}
                d, m, y = today.day, today.month, today.year
                run.text = run.text.replace("DATE", f"«{d}» {months[m]} {y}")
            if "CUSTOMER" in run.text:
                if customer.payment_type == "self":
                    run.text = run.text.replace("CUSTOMER", user.name)
                else:
                    run.text = run.text.replace("CUSTOMER", customer.name)
            if "FACE" in run.text:
                if customer.payment_type == "juridical":
                    run.text = run.text.replace("FACE", customer.behalf_genitive)
                else:
                    run.text = run.text.replace("FACE", "-")
            if "DOCUMENT" in run.text:
                if customer.payment_type == "juridical":
                    run.text = run.text.replace(
                        "DOCUMENT",
                        f" {customer.document} от " f"{customer.document_date.strftime('%d.%m.%Y')}",
                    )
                else:
                    run.text = run.text.replace("DOCUMENT", "")
            if "STUDENTA" in run.text:
                if customer.payment_type == "self":
                    run.text = run.text.replace("STUDENTA", "он же")
                else:
                    run.text = run.text.replace("STUDENTA", user.name)
            if "FORM" in run.text:
                if user.degree == "bachelor":
                    run.text = run.text.replace("FORM", "бакалавриата")
                else:
                    run.text = run.text.replace("FORM", "магистратуры")
            if "DIRECTION" in run.text:
                run.text = run.text.replace("DIRECTION", user.direction)
            if "PROGRAM" in run.text:
                run.text = run.text.replace("PROGRAM", "«" + user.program + "»")
            if "PERIOD" in run.text:
                if user.degree == "bachelor":
                    run.text = run.text.replace("PERIOD", "4")
                else:
                    run.text = run.text.replace("PERIOD", "2")
            if "DEGREE" in run.text:
                if user.degree == "bachelor":
                    run.text = run.text.replace("DEGREE", "бакалавра")
                else:
                    run.text = run.text.replace("DEGREE", "магистра")
            if "TOTAL" in run.text:
                run.text = run.text.replace("TOTAL", total[user.program])
            if "ANNUAL" in run.text:
                run.text = run.text.replace("ANNUAL", annual[user.program])
            if "PASSPORTCUST" in run.text:
                if customer.payment_type == "juridical":
                    run.text = run.text.replace(
                        "PASSPORTCUST",
                        "адрес: {0}, ОГРН {1}, ИНН {2}\n\n{3}".format(
                            customer.address, customer.ogrn, customer.inn, customer.behalf
                        ),
                    )
                elif customer.payment_type == "individual":
                    run.text = run.text.replace(
                        "PASSPORTCUST",
                        "Паспорт гражданина РФ {0}, выдан {1}, адрес: {2}, "
                        "{3}\n\n{4}".format(
                            customer.document.number,
                            customer.document.given_by,
                            customer.address,
                            customer.phone,
                            customer.name,
                        ),
                    )
                else:
                    run.text = run.text.replace(
                        "PASSPORTCUST",
                        "Паспорт гражданина РФ {0}, выдан {1} {2}, адрес: {3}, "
                        "{4}, {5}\n\n{6}".format(
                            user.document.number,
                            user.document.given_by,
                            user.document.issue_date.strftime("%d.%m.%Y"),
                            user.address,
                            user.phone,
                            user.email,
                            user.name,
                        ),
                    )
            if "STUDENTB" in run.text:
                if customer.payment_type == "self":
                    run.text = run.text.replace("STUDENTB", "См. графу «ЗАКАЗЧИК»")
                else:
                    run.text = run.text.replace(
                        "STUDENTB",
                        "{0}; Паспорт гражданина РФ {1}, выдан {2} {3}, адрес: "
                        "{4}, {5}, {6}\n\n{7}".format(
                            user.name,
                            user.document.number,
                            user.document.given_by,
                            user.document.issue_date.strftime("%d.%m.%Y"),
                            user.address,
                            user.phone,
                            user.email,
                            user.name,
                        ),
                    )
            if "Отметка о согласии на заключение настоящего Договора" in run.text:
                end_flag = True
            if user_age >= 18 and end_flag:
                run.text = ""

    contract.save("/".join(contract_path.split("/")[:-1]) + "/output_contract.docx")


def fill_receipt(user, customer, semesters, receipt_path="../logic/data/sample_receipt.docx"):
    annual_amount = {
        "Бизнес-информатика": 301000,
        "Технологии и инновации": 301000,
        "Управление высокотехнологичным бизнесом": 349000,
        "Технологии и стратегии бизнес-трансформации": 349000,
        "Цифровые продукты: создание и управление": 380000,
        "Стратегическое управление интеллектуальной собственностью / IP Management Strategy": 349000,
    }
    receipt = Document(receipt_path)

    for row in receipt.tables[0].rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if "CUSTOMER" in run.text:
                        if customer.payment_type == "self":
                            run.text = run.text.replace("CUSTOMER", user.name)
                        else:
                            run.text = run.text.replace("CUSTOMER", customer.name)
                    if "STUDENT" in run.text:
                        run.text = run.text.replace("STUDENT", user.name)
                    if "SEMESTER" in run.text:
                        if int(semesters) == 1:
                            run.text = run.text.replace("SEMESTER", "1 семестр")
                        else:
                            run.text = run.text.replace("SEMESTER", "1,2 семестры")
                    if "SUM" in run.text:
                        if int(semesters) == 1:
                            run.text = run.text.replace("SUM", str(int(0.55 * annual_amount[user.program])))
                        else:
                            run.text = run.text.replace("SUM", str(annual_amount[user.program]))

    receipt.save("/".join(receipt_path.split("/")[:-1]) + "/output_receipt.docx")


def autofill(cookies, current_url, semesters):
    user, customer = get_data(cookies, current_url)
    fill_contract(user, customer)
    fill_receipt(user, customer, semesters)


if __name__ == "__main__":
    cookie = "__ddg1_=DnvGJxZGJS1b8Z6JGV9R; _ga_K1E3SJW9RB=GS1.1.1687698828.7.0.1687698837.0.0.0; _ym_uid=1644441470443176458; _ym_d=1687790885; _gcl_au=1.1.376710900.1687795991; _ga_VBCPPQG8SR=GS1.1.1687944787.1.1.1687944800.0.0.0; _ga_6WNL1RBXB9=GS1.2.1689839868.6.0.1689839868.0.0.0; _ga_K8R8VFV12P=GS1.2.1689839868.6.0.1689839868.60.0.0; _ga=GA1.1.926234014.1687796086; _ga_EG4H99E35G=GS1.1.1689950977.31.1.1689951063.0.0.0; _ym_isad=2; access_token=eyJhbGciOiJSUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICIwSVliSmNVLW1wbEdBdzhFMzNSNkNKTUdWa3hZdUQ2eUItdWt3RlBJOXV3In0.eyJleHAiOjE2OTAxOTkyNjcsImlhdCI6MTY5MDE5NzQ2NywiYXV0aF90aW1lIjoxNjg2OTE1MTQ4LCJqdGkiOiJmNWQ5YjA2Ni02ZWE0LTQyMmQtYTIxYy0zMWY1MTIzYThjNGQiLCJpc3MiOiJodHRwczovL2lkLml0bW8ucnUvYXV0aC9yZWFsbXMvaXRtbyIsInN1YiI6IjU5YjY1MDI0LTE4NWItNDMxYS1hOWRlLWUzYTM4MTA5ZTAzOSIsInR5cCI6IkJlYXJlciIsImF6cCI6ImFiaXQtbGsiLCJzZXNzaW9uX3N0YXRlIjoiMjI2Y2Q0OGYtNjc3Ni00ZDI2LWEzODYtNDkxZmJmY2Y0ZGIzIiwiYWNyIjoiMCIsImFsbG93ZWQtb3JpZ2lucyI6WyJodHRwczovL2FiaXRsay5pdG1vLnJ1Il0sInNjb3BlIjoib3BlbmlkIGVtYWlsIHBob25lIHByb2ZpbGUiLCJzaWQiOiIyMjZjZDQ4Zi02Nzc2LTRkMjYtYTM4Ni00OTFmYmZjZjRkYjMiLCJpc3UiOjMxMjY2NSwicHJlZmVycmVkX3VzZXJuYW1lIjoiZHB0Z28iLCJlbWFpbCI6ImRuX3N0QGJrLnJ1In0.AELXdUe5rFrRVqrgqJEJuzm0XHj1rWf5TP5WgkrzYgZO93qLcGhXNWY1SZQk3_5e38L8FtPBlyJIraa4UvSqf3Afg-nmtbfxu1jNDlndjvWQ_xJlxELH8hKv_3OF4C_-0xrOGGekhidxE0MH5t-WjQSCucA7nkYCXbwECtgd09U8_JSgDTwksNMawtArvQkxwmpiTgfDVwQFqvDWoYvw9olgHh-AiLh4gp1a_7ciom3z4Mo1OJ8Y24qT11IqGbQuODxZwSbv9Uf_tn_A3xVT3bDIhOzO7KVpG3avNirrFq5sYAreA6I1nYfn5uj8Iq5hq3FHLFQPeJirDbAI4c-PHg; refresh_token=eyJhbGciOiJIUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICJjNTcxMzVjYy01ZjEwLTQ4ZTAtYTU5ZS1lYTYwMmY3ZTcxYzAifQ.eyJleHAiOjE2OTI3ODk0NjcsImlhdCI6MTY5MDE5NzQ2NywianRpIjoiZjA4ODgyYjMtZTE1MC00YjZhLWFkOGEtNDRiOWIzM2NjZTBlIiwiaXNzIjoiaHR0cHM6Ly9pZC5pdG1vLnJ1L2F1dGgvcmVhbG1zL2l0bW8iLCJhdWQiOiJodHRwczovL2lkLml0bW8ucnUvYXV0aC9yZWFsbXMvaXRtbyIsInN1YiI6IjU5YjY1MDI0LTE4NWItNDMxYS1hOWRlLWUzYTM4MTA5ZTAzOSIsInR5cCI6IlJlZnJlc2giLCJhenAiOiJhYml0LWxrIiwic2Vzc2lvbl9zdGF0ZSI6IjIyNmNkNDhmLTY3NzYtNGQyNi1hMzg2LTQ5MWZiZmNmNGRiMyIsInNjb3BlIjoib3BlbmlkIGVtYWlsIHBob25lIHByb2ZpbGUiLCJzaWQiOiIyMjZjZDQ4Zi02Nzc2LTRkMjYtYTM4Ni00OTFmYmZjZjRkYjMifQ.Q2ofhwMjbnY1Cte1CSYtYT25Y4WkQQNJS4KGPJVK_-E; _ym_visorc=b"
    url = "https://abitlk.itmo.ru/window/0/questionnaire/29985/45949/programs"

    autofill(cookie, url)
