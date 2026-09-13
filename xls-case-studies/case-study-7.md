# 📊 Case Study 07 — Dynamic KPI Reporting & Lookup Dashboard

## 🎯 Cíl projektu

Cílem této mini case study bylo procvičit:

- Excel Tables (`Ctrl + T`)
- strukturované odkazy (Structured References)
- funkce `INDEX()` a `POZVYHLEDAT()`
- funkci `IFS()` a business logiku
- funkce `SUBTOTAL()` a `AGGREGATE()`
- funkci `SOUČIN.SKALÁRNÍ()`
- dynamické oblasti pomocí funkce `POSUN()`
- parametrické výpočty
- tvorbu jednoduchého dynamického dashboardu

---

# 📁 Dataset

| ID | Produkt | Kategorie | Region | Tržba | Náklady |
|---:|---|---|---|---:|---:|
| 101 | Notebook | Elektronika | Praha | 85000 | 62000 |
| 102 | Monitor | Elektronika | Brno | 42000 | 31000 |
| 103 | Myš | Příslušenství | Ostrava | 12000 | 5000 |
| 104 | Klávesnice | Příslušenství | Praha | 18000 | 9000 |
| 105 | Tiskárna | Kancelář | Brno | 35000 | 28000 |
| 106 | Router | Síť | Ostrava | 26000 | 15000 |
| 107 | SSD Disk | Komponenty | Praha | 47000 | 32000 |
| 108 | Webkamera | Příslušenství | Brno | 16000 | 7000 |

---

# ✅ Business pravidla a výpočty

## 1️⃣ Výpočet marže

Vzorec:

```excel
=([@Tržba]-[@Náklady])/[@Tržba]
```

Výsledek:

- výpočet procentuální marže
- použití strukturovaných odkazů
- automatické rozkopírování vzorce v Excel Table

---

## 2️⃣ Business status pomocí `IFS()`

Business pravidla:

- **TOP** → marže je vyšší než 30 % a zároveň tržba vyšší než 40 000
- **RIZIKOVÝ** → marže je nižší než 20 %
- **STANDARD** → všechny ostatní případy

Vzorec:

```excel
=IFS(A([@Marže]>0,3;[@Tržba]>40000);"TOP";[@Marže]<0,2;"RIZIKOVÝ";PRAVDA;"STANDARD")
```

Výsledek:

| Produkt | Status |
|---|---|
| SSD Disk | TOP |
| Ostatní produkty | STANDARD |

---

## 3️⃣ Dynamický produktový lookup

Vybraný produkt:

```text
Notebook
```

Použité funkce:

- `INDEX()`
- `POZVYHLEDAT()`

Příklad vzorce pro vrácení regionu:

```excel
=INDEX(Tabulka1[Region];POZVYHLEDAT($K$2;Tabulka1[Produkt];0))
```

Dashboard podle zvoleného produktu automaticky vrací:

- kategorii
- region
- tržbu
- náklady
- marži
- status

---

## 4️⃣ Viditelné tržby pomocí `SUBTOTAL()`

Vzorec:

```excel
=SUBTOTAL(109;Tabulka1[Tržba])
```

Význam:

- `9` = součet filtrovaných dat, ale zahrnuje ručně skryté řádky
- `109` = součet, který ignoruje filtrované i ručně skryté řádky

Použití:

- filtrování reportů
- dynamické KPI
- interaktivní reporting

---

## 5️⃣ Počet TOP produktů pomocí `SOUČIN.SKALÁRNÍ()`

Vzorec:

```excel
=SOUČIN.SKALÁRNÍ((Tabulka1[Status]="TOP")*1)
```

Význam:

- logické hodnoty `PRAVDA` a `NEPRAVDA` se převedou na `1` a `0`
- produkty lze spočítat podle podmínky bez dalšího pomocného sloupce

Výsledek:

```text
1
```

---

## 6️⃣ Průměr marže pomocí `AGGREGATE()`

Vzorec:

```excel
=AGGREGATE(1;7;Tabulka1[Marže])
```

Význam:

- `1` = průměr
- `7` = ignorování skrytých řádků a chybových hodnot

Použití:

- robustní reporting
- práce s filtrovanými daty
- KPI monitoring

---

## 7️⃣ Dynamický součet posledních X položek

Parametr:

```text
Počet posledních položek = 3
```

Vzorec:

```excel
=SUMA(POSUN(Tabulka1[[#Záhlaví];[Tržba]];POČET2(Tabulka1[Tržba])-K18+1;0;K18;1))
```

Výsledek:

```text
89000
```

Význam:

- vytvoření dynamické oblasti
- parametrický výpočet podle zadaného počtu položek
- automatické přizpůsobení při přidání dalších dat

---

# 🧠 Co jsem se naučil

- `INDEX()` a `POZVYHLEDAT()` představují flexibilnější alternativu k funkci `SVYHLEDAT()`
- strukturované odkazy zlepšují čitelnost vzorců
- `SUBTOTAL()` reaguje na filtrování tabulky
- `AGGREGATE()` dokáže ignorovat chyby i skryté řádky
- `SOUČIN.SKALÁRNÍ()` umožňuje provádět logické výpočty bez pomocných sloupců
- `POSUN()` umožňuje vytvářet dynamické oblasti
- Excel může fungovat jako jednoduchý interaktivní analytický dashboard

---

# 🛠 Použité nástroje

- Excel Tables (`Ctrl + T`)
- Structured References
- `INDEX()`
- `POZVYHLEDAT()`
- `IFS()`
- `A()`
- `SUBTOTAL()`
- `AGGREGATE()`
- `SOUČIN.SKALÁRNÍ()`
- `POSUN()`
- `POČET2()`

---

# 🔗 BI mindset

Tato case study simuluje jednoduchý analytický dashboard s dynamickými KPI a produktovým detailem.

Podobný princip se používá v:

- sales reportingu
- KPI monitoringu
- controllingu
- Power BI dashboardech
- business reportingu
- interaktivních filtrech

---

# 🚀 Výsledek

Pomocí pokročilých funkcí Excelu jsem vytvořil:

- dynamický produktový lookup
- business klasifikaci produktů
- KPI reporting reagující na filtrování
- robustní agregace
- parametrický výpočet posledních X položek
- dynamickou práci s rozšiřujícími se daty

Case study mi umožnila procvičit workflow podobné reálné práci juniorního datového analytika.
