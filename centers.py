import json
import os

from app_paths import config_path, resolve_template


DEFAULT_CENTERS = [
    {
        "id": "Novoros",
        "button": "Новороссийск",
        "name": "Новороссийск",
        "template": "Novoros.docx",
        "template_fizo": "",
        "east": False,
        "workers": [
            "Лагутин Ростислав Борисович",
            "Соловьев Сергей Викторович",
            "Соловьев Артем Андреевич",
        ],
    },
    {
        "id": "Astrakhan",
        "button": "Астрахань",
        "name": "Астрахань",
        "template": "Astrakhan.docx",
        "template_fizo": "Astrakhan_FIZO.docx",
        "east": False,
        "workers": ["Шматова Раиса Анатольевна", "Сивоконь Светлана Викторовна"],
    },
    {
        "id": "Bor",
        "button": "Бор",
        "name": "Бор",
        "template": "Bor.docx",
        "template_fizo": "Bor_FIZO.docx",
        "east": False,
        "workers": ["Селюнин Александр Николаевич", "Аладьин Алексей Николаевич"],
    },
    {
        "id": "Ekaterinburg",
        "button": "Екатеринбург",
        "name": "Екатеринбург",
        "template": "Ekaterinburg.docx",
        "template_fizo": "",
        "east": True,
        "workers": ["Рослякова Надежда Викторовна"],
    },
    {
        "id": "Habarovsk",
        "button": "Хабаровск",
        "name": "Хабаровск",
        "template": "Habarovsk.docx",
        "template_fizo": "Habarovsk_FIZO.docx",
        "east": True,
        "workers": ["Смирных Евгений Михайлович", "Смирных Надежда Евгеньевна"],
    },
    {
        "id": "Kaliningrad",
        "button": "Калининград",
        "name": "Калининград",
        "template": "Kaliningrad.docx",
        "template_fizo": "Kaliningrad_FIZO.docx",
        "east": False,
        "workers": ["Дьяков Александр Евгеньевич"],
    },
    {
        "id": "NN",
        "button": "Нижний Новгород",
        "name": "Нижний Новгород",
        "template": "NN.docx",
        "template_fizo": "",
        "east": False,
        "workers": ["Еркович Наталья Евгеньевна"],
    },
    {
        "id": "Murmansk",
        "button": "Мурманск",
        "name": "Мурманск",
        "template": "Murmansk.docx",
        "template_fizo": "Murmansk_FIZO.docx",
        "east": False,
        "workers": ["Пахомов Андрей Юрьевич", "Лазарева Валерия Андреевна"],
    },
    {
        "id": "Nahodka",
        "button": "Находка",
        "name": "Находка",
        "template": "Nahodka.docx",
        "template_fizo": "",
        "east": True,
        "workers": ["Равнянский Константин Витальевич", "Зубакин Эдуард Николаевич"],
    },
    {
        "id": "PK",
        "button": "Петропавловск-Камчатский",
        "name": "Петропавловск-Камчатский",
        "template": "PK.docx",
        "template_fizo": "PK_FIZO.docx",
        "east": True,
        "workers": [
            "Каушанов Александр Викторович",
            "Еремеев Геннадий Гайсович",
            "Зотов Станислав Викторович",
        ],
    },
    {
        "id": "Rostov",
        "button": "Ростов-на-Дону",
        "name": "Ростов",
        "template": "Rostov.docx",
        "template_fizo": "",
        "east": False,
        "workers": ["Евстратова Наталья Павловна", "Язьков Анатолий Сергеевич"],
    },
    {
        "id": "Samara",
        "button": "Самара",
        "name": "Самара",
        "template": "Samara.docx",
        "template_fizo": "",
        "east": False,
        "workers": ["Башмакова Виктория Владимировна"],
    },
    {
        "id": "Sevastopol",
        "button": "Севастополь",
        "name": "Севастополь",
        "template": "Sevastopol.docx",
        "template_fizo": "",
        "east": False,
        "workers": ["Кошелев Тимур Сергеевич", "Кошелева Джеваире Айдеровна"],
    },
    {
        "id": "Spb",
        "button": "Санкт-Петербург",
        "name": "Спб",
        "template": "Spb.docx",
        "template_fizo": "",
        "east": False,
        "workers": ["Савельев Дмитрий Генрихович", "Семенова Мария Сергеевна"],
    },
    {
        "id": "US",
        "button": "Южно-Сахалинск",
        "name": "Сахалин",
        "template": "US.docx",
        "template_fizo": "US_FIZO.docx",
        "east": True,
        "workers": ["Мельчарик Владимир Юрьевич", "Шадрин Владимир Валерьевич"],
    },
    {
        "id": "Vladivostok",
        "button": "Владивосток",
        "name": "Владивосток",
        "template": "Vladivostok.docx",
        "template_fizo": "Vladivostok_FIZO.docx",
        "east": True,
        "workers": [
            "Проценко Игорь Владимирович",
            "Лабонин Илья Валентинович",
            "Родионова Елена Ивановна",
        ],
    },
    {
        "id": "Krasnoyarsk",
        "button": "Красноярск",
        "name": "Красноярск",
        "template": "Krasnoyarsk.docx",
        "template_fizo": "Krasnoyarsk_FIZO.docx",
        "east": True,
        "workers": ["Бражевская Елена Викторовна"],
    },
]


def load_centers():
    path = config_path()
    if not os.path.exists(path):
        save_centers(DEFAULT_CENTERS)
        return [dict(center) for center in DEFAULT_CENTERS]

    with open(path, "r", encoding="utf-8") as file:
        centers = json.load(file)
    return centers


def save_centers(centers):
    with open(config_path(), "w", encoding="utf-8") as file:
        json.dump(centers, file, ensure_ascii=False, indent=2)


def normalized_center(center):
    normalized = dict(center)
    normalized["template_path"] = resolve_template(center.get("template", ""))
    normalized["template_fizo_path"] = resolve_template(center.get("template_fizo", ""))
    normalized["workers"] = list(center.get("workers", []))
    normalized["east"] = bool(center.get("east", False))
    return normalized
