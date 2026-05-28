import json
import os

from app_paths import people_config_path


DEFAULT_PEOPLE = {
    "chiefs": [
        {
            "id": "Беззубцев",
            "label": "Беззубцев А.А.",
            "job_title": "Начальник отдела проверок сил ОТБ",
            "fullname": "А.А. Беззубцев",
        },
        {
            "id": "Алябьев",
            "label": "Алябьев А.Б.",
            "job_title": "Заместитель начальника Службы",
            "fullname": "А.Б. Алябьев",
        },
        {
            "id": "Попырина",
            "label": "Попырина Е.М.",
            "job_title": "Заместитель начальника отдела проверок сил ОТБ",
            "fullname": "Е.М. Попырина",
        },
    ],
    "issuers": [
        {
            "id": "Маруся",
            "label": "Садовникова Ю.В.",
            "fullname": "Садовникова Юлия Владимировна",
            "email": "oa18@msecurity.ru",
        },
        {
            "id": "Вантуз",
            "label": "Воротилов И.И.",
            "fullname": "Воротилов Иван Иванович",
            "email": "oa13@msecurity.ru",
        },
        {
            "id": "Создатель!!!",
            "label": "Желудков А.В.",
            "fullname": "Желудков Андрей Викторович",
            "email": "oa15@msecurity.ru",
        },
        {
            "id": "Максим",
            "label": "Миронов М.С.",
            "fullname": "Миронов Максим Сергеевич",
            "email": "oa6@msecurity.ru",
        },
        {
            "id": "Катя",
            "label": "Попырина Е.М.",
            "fullname": "Попырина Екатерина Михайловна",
            "email": "oa3@msecurity.ru",
        },
        {
            "id": "АА",
            "label": "Беззубцев А.А.",
            "fullname": "Беззубцев Александр Анатольевич",
            "email": "oa2@msecurity.ru",
        },
    ],
}


def load_people():
    path = people_config_path()
    if not os.path.exists(path):
        save_people(DEFAULT_PEOPLE)
        return _copy_people(DEFAULT_PEOPLE)

    with open(path, "r", encoding="utf-8") as file:
        people = json.load(file)

    people.setdefault("chiefs", [])
    people.setdefault("issuers", [])
    return people


def save_people(people):
    with open(people_config_path(), "w", encoding="utf-8") as file:
        json.dump(people, file, ensure_ascii=False, indent=2)


def _copy_people(people):
    return {
        "chiefs": [dict(item) for item in people.get("chiefs", [])],
        "issuers": [dict(item) for item in people.get("issuers", [])],
    }
