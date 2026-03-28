# 📦 Wiekowanie zapasów i kalkulacja rezerw

Aplikacja Streamlit do automatycznego wiekowania zapasów, mapowania typów materiałów i kalkulacji rezerw bilansowych.

---

## 🚀 Szybki start

### Wymagania
- Python 3.11+
- pip

### Instalacja lokalna

```bash
# 1. Sklonuj lub rozpakuj projekt
cd inventory_app

# 2. Utwórz wirtualne środowisko (zalecane)
python -m venv venv
source venv/bin/activate        # Linux/macOS
venv\Scripts\activate           # Windows

# 3. Zainstaluj zależności
pip install -r requirements.txt

# 4. Uruchom aplikację
streamlit run app.py
```

Aplikacja otworzy się automatycznie pod adresem: `http://localhost:8501`

---

## 🌐 Wdrożenie online – Streamlit Community Cloud (bezpłatne)

1. Utwórz konto na [share.streamlit.io](https://share.streamlit.io)
2. Wgraj projekt na **GitHub** (w tym plik `data/default_mapping.xlsx`)
3. Kliknij **"New app"** → wskaż repozytorium → plik główny: `app.py`
4. Kliknij **Deploy**

> ⚠️ Upewnij się, że plik `data/default_mapping.xlsx` jest w repozytorium (nie w `.gitignore`).

---

## 🌐 Wdrożenie online – Render.com

1. Utwórz konto na [render.com](https://render.com)
2. Nowy **Web Service** → połącz z repozytorium GitHub
3. Ustaw:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
4. Kliknij **Deploy**

---

## 📁 Struktura projektu

```
inventory_app/
├── app.py                  # Główna aplikacja Streamlit
├── processing.py           # Logika biznesowa (wiekowanie, mapowania, rezerwy)
├── export.py               # Eksport do Excel/CSV
├── utils.py                # Funkcje pomocnicze, formatowanie
├── requirements.txt        # Zależności Python
├── README.md               # Ta instrukcja
└── data/
    └── default_mapping.xlsx  # Domyślny plik mappingów (Mapp1 + Mapp2)
```

---

## 🗂️ Jak zmienić domyślny mapping

1. Przygotuj plik Excel z dwoma arkuszami:
   - **Mapp1** – lista indeksów PROWAX w kolumnie B (nagłówek w wierszu 1)
   - **Mapp2** – tabela z nagłówkami: `Type of materials`, `Magazyn`, `Typ surowca`
2. Zastąp plik `data/default_mapping.xlsx` swoim plikiem
3. Uruchom ponownie aplikację

---

## 📊 Logika biznesowa

### Mapping 1 – Rodzaj indeksu
| Warunek | Wynik |
|---------|-------|
| Indeks w Mapp1 | PROWAX |
| Indeks nieobecny w Mapp1 | NON PROWAX |

### Mapping 2 – Type of materials
Mapowanie po kombinacji `Magazyn + Typ surowca` z arkusza Mapp2.
Brak mapowania → `UNMAPPED` (rezerwa = 0%).

### Wiekowanie
| Przedział | Opis |
|-----------|------|
| 0-3 mcy | 0–2 pełne miesiące |
| 3-6 mcy | 3–5 miesięcy |
| 6-9 mcy | 6–8 miesięcy |
| 9-12 mcy | 9–11 miesięcy |
| pow 12 mcy | 12+ miesięcy |

### % rezerwy
| Type | 0-3 | 3-6 | 6-9 | 9-12 | >12 |
|------|-----|-----|-----|------|-----|
| RW   | 0%  | 0%  | 0%  | 50%  | 100% |
| WIP  | 0%  | 50% | 100%| 100% | 100% |
| FG   | 0%  | 0%  | 100%| 100% | 100% |

### Status nowa/nabyta
| Rodzaj | Granica |
|--------|---------|
| PROWAX | 2023-11-01 |
| NON PROWAX | 2023-12-01 |

---

## 🤝 Założenia przyjęte w kodzie

- Kolumna `Wartość mag.` parsowana do liczby; błędne wartości → 0 PLN
- Indeksy PROWAX porównywane case-insensitive po normalizacji do string
- Nagłówki Mapp2 wyszukiwane dynamicznie (nie na sztywno po numerze wiersza)
- Type `x` w Mapp2 traktowany jak UNMAPPED (rezerwa 0%)
- Data przyjęcia > data analizy → bucket `data > dzień analizy`, rezerwa 0%
- Puste daty → bucket `błąd daty`, rezerwa 0%
