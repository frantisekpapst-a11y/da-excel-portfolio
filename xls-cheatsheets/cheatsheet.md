# 📊 Excel Advanced Cheatsheet

## Klávesové zkratky • Vzorce • Lookupy • Filtry • Agregace • BI mindset

> 🎯 Praktický cheatsheet ke kurzu Excel pro pokročilé  
> Zaměřeno na efektivní workflow datového analytika.

---

# 🚀 1. Klávesové zkratky
| Zkratka | Funkce |
|---|---|
| `Ctrl + Home` | Přesun na A1 |
| `Ctrl + End` | Poslední použitá buňka |
| `Ctrl + PgDown` | Další list |
| `Ctrl + PgUp` | Předchozí list |
| `Ctrl + šipka` | Přesun na konec dat |
| `Ctrl + Shift + šipka` | Rozšíření výběru |
| `Ctrl + Space` | Celý sloupec |
| `Shift + Space` | Celý řádek |
| `Ctrl + A` | Výběr tabulky |
| `Ctrl + A` 2× | Výběr celého listu |
| `Ctrl + ;` | Aktuální datum |
| `F2` | Editace buňky |
| `F4` | Přepínání adresace |
| `Ctrl + T` | Excel tabulka |
| `Ctrl + Enter` | Hromadné vyplnění |
| `Ctrl + Shift + L` | Zapnout/vypnout filtr |
| `Ctrl + Alt + V` | Vložit jinak |
| `Ctrl + F3` | Správce názvů |
| `Alt + F11` | VBA editor |

---

# 📐 2. Adresace buněk
| Typ | Příklad | Chování |
|---|---|---|
| Relativní | `A1` | mění řádek i sloupec |
| Absolutní | `$A$1` | nemění nic |
| Smíšená | `$A1` | fixní sloupec |
| Smíšená | `A$1` | fixní řádek |

## F4 cyklus

```excel
A1 → $A$1 → A$1 → $A1 → A1
Absolutní adresa

Použití:

DPH
kurz měny
sleva
target
bonus %
centrální parametr
=A2*$B$2

Lepší varianta:

=A2*sleva

🧮 3. Smíšené adresování
=$A2*B$1
Část	Význam
$A2	fixní sloupec A
B$1	fixní řádek 1

Použití:

sazebníky
matice cen
forecasty
scénáře

📊 4. Rozšiřování vzorců
Metoda	Co kopíruje
Fill Handle	vzorec + formát
Ctrl + Enter	jen obsah / vzorec
označit oblast
→ napsat vzorec
→ Ctrl + Enter

Výhody:

zachování formátu
rychlost
menší chybovost

📋 5. Excel tabulky (Ctrl + T)

Výhody:

automatické rozkopírování vzorců
dynamické rozsahy
filtry
structured references
lepší kompatibilita s Power Query / Power BI
Structured references
=[@Ks]*[@Cena]
=KDYŽ([@Tržba]>20000;"HIGH";"STANDARD")

🧠 6. Pojmenované oblasti

Místo:

=SUMA(B2:B5000)

lepší:

=SUMA(trzby_q1)
Zkratka	Funkce
Ctrl + F3	Správce názvů
F3	Vložení názvu

📑 7. Listy a viditelnost
Typ	Použití
Datový list	data, výpočty, dashboardy
Grafický list	jeden graf
Skrytý list	lze zobrazit v Excelu
Very Hidden	pouze přes VBA

🏗️ 8. Hromadná práce s listy
Akce	Postup
Nesousedící listy	Ctrl + klik
Sousedící listy	Shift + klik
Všechny listy	pravé tlačítko → Vybrat všechny listy
3D reference
=SUMA('Leden:Prosinec'!B2)

Použití:

měsíční reporting
regionální reporting
pobočky

🧩 9. Vnořování funkcí
=ZAOKROUHLIT(PRŮMĚR(A1:A5);0)

Excel počítá:

nejdřív PRŮMĚR()
potom ZAOKROUHLIT()
KDYŽ()
=KDYŽ(B2>10000;"HIGH";"LOW")
IFS()
=IFS(B2<10000;"LOW";B2<50000;"MEDIUM";B2>=50000;"HIGH")

⚖️ 10. Operátory
Operátor	Význam
>	větší než
<	menší než
=	rovná se
>=	větší nebo rovno
<=	menší nebo rovno
<>	nerovná se
Textový operátor
=A1&" "&B1

🔍 11. Zástupné znaky
Znak	Význam
*	libovolný počet znaků
?	právě jeden znak
=COUNTIF(A:A;"Praha*")

📊 12. Maticové vzorce
=B2:B6*C2:C6

Moderní Excel 365 výsledek automaticky rozlije do více buněk.

🔄 13. TRANSPOZICE
=TRANSPOZICE(A1:A5)

Použití:

řádek → sloupec
sloupec → řádek

📈 14. Kumulativní výpočty
=SUMA($C$3:C3)
Část	Význam
$C$3	fixní začátek
C3	pohyblivý konec

Použití:

running totals
YTD
cash flow
budget tracking

🔍 15. SVYHLEDAT
=SVYHLEDAT(hledat;tabulka;sloupec;typ)
Přesná shoda
=SVYHLEDAT(A2;$A$2:$B$10;2;0)
Přibližná shoda
=SVYHLEDAT(A2;$A$2:$B$10;2;1)

Použití:

ceníky
sazebníky
intervaly
Limity
hledá jen doprava
pevné číslo sloupce
může se rozbít při vložení sloupce

🔎 16. POZVYHLEDAT
=POZVYHLEDAT("Cena";A1:E1;0)

Vrací pozici hodnoty.

Hodnota	Význam
0	přesná shoda
1	nejbližší menší
-1	nejbližší větší

🔗 17. INDEX + POZVYHLEDAT
=INDEX(B2:B10;POZVYHLEDAT(E2;A2:A10;0))

Výhody:

flexibilnější než SVYHLEDAT
lookup doleva
stabilnější při změně pořadí sloupců

🚀 18. XLOOKUP
=XLOOKUP(co;kde_hledat;co_vratit)
=XLOOKUP(G2;A2:A10;E2:E10;"nenalezeno")
Výhody
hledá vlevo i vpravo
není potřeba číslo sloupce
umí náhradní hodnotu při nenalezení
stabilnější než SVYHLEDAT
Režimy shody
Hodnota	Význam
0	přesná shoda
-1	přesná nebo nejbližší menší
1	přesná nebo nejbližší větší
2	wildcard *, ?

🚨 19. IFERROR
=IFERROR(A1/B1;"chyba")

Použití:

ošetření lookupů
stabilní dashboardy
čisté reporty

🔎 20. Jednoduchý filtr
Zkratka	Funkce
Ctrl + Shift + L	zapnutí/vypnutí filtru
Logika
Typ	Význam
A	musí platit obě podmínky
NEBO	stačí jedna podmínka
Kaskádové filtrování
další filtr pracuje jen nad viditelnými řádky

🧠 21. Rozšířený filtr

Rozšířený filtr:

používá oblast kritérií
zvládá komplexní logiku
umí exportovat výsledky
podporuje výpočtová kritéria
Rozložení kritérií	Význam
vedle sebe	AND
pod sebou	OR
Příklad AND
Věk	Bydliště
<40	Praha
Příklad OR
Bydliště
Praha
Brno

🧮 22. Výpočtová kritéria
Nadprůměrná hodnota
=$E2>PRŮMĚR($E$2:$E$13)
Nejvyšší hodnota
=$C2=MAX($C$2:$C$13)
Nejnižší hodnota
=$C2=MIN($C$2:$C$13)

🧩 23. Průřezy

Průřezy fungují nad:

Excel tabulkou
kontingenční tabulkou
Ctrl + T
→ Návrh tabulky
→ Vložit průřez

Výhody:

vizuální filtrování
interaktivní dashboard
rychlá analýza

🧮 24. Podmíněné agregační funkce – jedna podmínka
COUNTIF

Počet záznamů podle podmínky.

=COUNTIF(oblast;kritérium)

Příklad:

=COUNTIF(A2:A11;"pondělí")
SUMIF

Součet hodnot podle podmínky.

=SUMIF(oblast;kritérium;oblast_součtu)

Příklad:

=SUMIF(A2:A11;"pondělí";B2:B11)
AVERAGEIF

Průměr hodnot podle podmínky.

=AVERAGEIF(oblast;kritérium;oblast_průměru)

Příklad:

=AVERAGEIF(A2:A11;"pondělí";B2:B11)
Vztah
AVERAGEIF = SUMIF / COUNTIF

🧮 25. Podmíněné agregační funkce – více kritérií
COUNTIFS

Počet řádků splňujících více podmínek.

=COUNTIFS(oblast1;kritérium1;oblast2;kritérium2)

Příklad:

=COUNTIFS(B2:B7;"Praha";C2:C7;"Notebook")
SUMIFS

Součet hodnot splňujících více podmínek.

=SUMIFS(oblast_součtu;oblast1;kritérium1;oblast2;kritérium2)

Příklad:

=SUMIFS(D2:D7;B2:B7;"Praha";C2:C7;"Notebook")
AVERAGEIFS

Průměr hodnot splňujících více podmínek.

=AVERAGEIFS(D2:D7;B2:B7;"Praha";C2:C7;"Notebook")
MINIFS / MAXIFS
=MINIFS(D2:D7;B2:B7;"Praha")
=MAXIFS(D2:D7;B2:B7;"Praha")

📌 26. Kritéria ve funkcích -IFS
Pevné kritérium
">4"
Kritérium s odkazem na buňku
">"&E3
Textový vzor
"p*"
Oboustranný interval
=SUMIFS(D2:D20;D2:D20;">"&E3;D2:D20;"<"&E4)

⚠️ 27. Limity SUMIFS / COUNTIFS

Funkce -IFS:

pracují primárně s logikou AND
neumí přímo složitou vlastní logiku
neumí sčítat více sloupců najednou v jedné oblasti součtu
někdy vyžadují pomocný sloupec
Alternativa
=SOUČIN.SKALÁRNÍ()

# 🧮 28. SOUČIN.SKALÁRNÍ – základy

`SOUČIN.SKALÁRNÍ()` násobí odpovídající hodnoty v oblastech a výsledky sečte.

```excel
=SOUČIN.SKALÁRNÍ(B2:B4;C2:C4)

Odpovídá logice:

B2*C2 + B3*C3 + B4*C4

Použití:

celková tržba
zisk
marže
vážený průměr
výpočty bez pomocných sloupců
Více činitelů
=SOUČIN.SKALÁRNÍ(Ks;Cena;Marže)
Vážený průměr
=SOUČIN.SKALÁRNÍ(Hodnoty;Váhy)/SUMA(Váhy)

🧮 29. SOUČIN.SKALÁRNÍ jako COUNTIF / SUMIF
Počet podle podmínky
=SOUČIN.SKALÁRNÍ((A2:A10="Praha")*1)

PRAVDA/NEPRAVDA se převede na 1/0.

Součet podle podmínky
=SOUČIN.SKALÁRNÍ((A2:A10="Praha")*(C2:C10))
Porovnání dvou sloupců
=SOUČIN.SKALÁRNÍ((Skutečnost>Plán)*1)

Výhoda:

není potřeba pomocný sloupec
umí porovnat hodnoty řádek po řádku

🧠 30. Logika ve skalárním součinu
AND logika
=SOUČIN.SKALÁRNÍ((A2:A10="Praha")*(B2:B10="Notebook")*(C2:C10))

* znamená:

A ZÁROVEŇ
OR logika
=SOUČIN.SKALÁRNÍ(((A2:A10="Praha")+(A2:A10="Brno")>0)*(C2:C10))

+ znamená:

NEBO

>0 zajistí, že se řádek splňující více OR podmínek započítá jen jednou.

Kombinace OR + AND
=SOUČIN.SKALÁRNÍ((((Kanál="Ads")+(Kanál="Organic"))>0)*(Marže>0,2)*(Vráceno=0)*(Tržba))

Čtení:

(Kanál = Ads NEBO Organic)
A ZÁROVEŇ
Marže > 20 %
A ZÁROVEŇ
Vráceno = 0

🧩 31. SOUČIN.SKALÁRNÍ nad maticí

Skalární součin může pracovat i s oblastí více řádků a sloupců.

Výběr řádků
=SOUČIN.SKALÁRNÍ((A3:A26<=6)*(B3:D26))

Sečte hodnoty z matice jen pro vybrané řádky.

Výběr sloupců
=SOUČIN.SKALÁRNÍ((B1:D1="CIO")*(B3:D26))

Sečte hodnoty jen z vybraných sloupců.

Výběr řádků i sloupců
=SOUČIN.SKALÁRNÍ((A3:A26<=6)*(B1:D1="CIO")*(B3:D26))
Vyhledání hodnoty v průsečíku
=SOUČIN.SKALÁRNÍ((A2:A5=G2)*(B1:E1=H2)*(B2:E5))

Použití:

matice vzdáleností
ceníky
sazebníky
výběr hodnoty podle řádku a sloupce

🗄️ 32. Databázové funkce

Databázové funkce pracují s celou tabulkou a samostatnou oblastí kritérií.

Syntaxe
=DFUNKCE(databáze;pole;kritéria)
Nejčastější funkce
Funkce	Význam
DSUMA()	podmíněný součet
DPRŮMĚR()	podmíněný průměr
DPOČET()	počet číselných hodnot
DPOČET2()	počet neprázdných hodnot
DMAX()	nejvyšší hodnota
DMIN()	nejnižší hodnota
Příklad DSUMA
=DSUMA(A1:C6;"Tržba";A8:B10)

nebo:

=DSUMA(A1:C6;3;A8:B10)
Kritéria
Rozložení	Význam
Vedle sebe	AND
Pod sebou	OR

Příklad:

Město	Produkt
Praha	A
Brno	A

Znamená:

(Praha A Produkt A)
NEBO
(Brno A Produkt A)

🧠 33. Logické funkce A() a NEBO()
A()
=A(B2>10000;C2>5)

Vrátí PRAVDA, pouze pokud platí všechny podmínky.

NEBO()
=NEBO(B2>10000;C2>5)

Vrátí PRAVDA, pokud platí alespoň jedna podmínka.

Kombinace s KDYŽ()
=KDYŽ(NEBO(A(B2>10000;C2>5);D2="ano");"VIP";"STANDARD")

Čtení:

VIP pokud:
(útrata > 10000 A objednávky > 5)
NEBO
premium = ano

🧩 34. Převod logiky na výstup
Logický výstup
=NEBO([@VIP]=1;[@Kredit]>0;[@Zaplaceno]>0)

Vrací:

PRAVDA / NEPRAVDA
Textový výstup
=KDYŽ(NEBO([@VIP]=1;[@Kredit]>0;[@Zaplaceno]>0);"ano";"ne")
Číselný výstup
=NEBO([@VIP]=1;[@Kredit]>0;[@Zaplaceno]>0)*1

Vrací:

1 / 0

🧠 35. Další logické funkce
Funkce	Význam
NE()	obrací logický výsledek
XOR()	exkluzivní NEBO
IFS()	více podmínek bez vnořeného KDYŽ
IFERROR()	ošetření chyby
KDYŽ()	převod logiky na výstup
XOR
=XOR(A1>0;B1>0)

Vrátí PRAVDA, pokud platí právě jedna z podmínek.

🎯 36. Business rule mindset

Při tvorbě pravidel se vždy ptej:

Musí platit všechno?
→ A()

Stačí jedna možnost?
→ NEBO()

Je potřeba výjimka?
→ NEBO() nebo vnořená logika

Potřebuji textový výsledek?
→ KDYŽ()

Potřebuji počítat výsledky?
→ *1

🔗 37. SQL → Excel → Python mindset
Koncept	SQL	Excel	Python / pandas
Filtr	WHERE	Filtr / KDYŽ()	boolean mask
AND	AND	A() / *	and / &
OR	OR	NEBO() / +	or / `
Podmínka	CASE WHEN	KDYŽ()	if / np.where()
Agregace	SUM()	SUMIFS()	groupby().sum()
Join	JOIN	XLOOKUP()	merge()
Předchozí hodnota	LAG()	posun řádku	shift()

⌨️ 38. Speciální znaky na CZ klávesnici
Znak	Zkratka
[	AltGr + F
]	AltGr + G
{	AltGr + B
}	AltGr + N
@	AltGr + V
$	AltGr + ů

🧱 39. Profesionální struktura analytického souboru
RAW_DATA
↓
TRANSFORM
↓
MODEL
↓
PIVOTS
↓
DASHBOARD
↓
EXPORT

🎯 40. SQL → Excel mindset
SQL	Excel
SELECT	výběr dat
WHERE	filtr / rozšířený filtr
GROUP BY	kontingenční tabulka
SUM()	SUMA() / SUMIF() / SUMIFS()
COUNT()	COUNTIF() / COUNTIFS()
AVG()	AVERAGEIF() / AVERAGEIFS()
JOIN	SVYHLEDAT() / XLOOKUP()
DISTINCT	Bez duplicitních záznamů

🎨 41. Podmíněné formátování pomocí vzorců
Zvýraznění celé řady
=$E2>10000

Použití:

označit celý řádek podle jedné hodnoty
KPI alerty
reporting
dashboardy

Princip:

označíš celou tabulku
vzorec se vyhodnocuje pro první řádek oblasti
relativní řádek = posouvá se
Nejvyšší / nejnižší hodnota
=$E2=MAX($E$2:$E$20)
=$E2=MIN($E$2:$E$20)

Použití:

best performer
nejhorší výsledek
top sales
lowest margin
Kombinace více podmínek
AND logika
=A($D2<0,2;$E2>3)

Význam:

marže < 20 %
A ZÁROVEŇ reklamace > 3
OR logika
=NEBO($E2=MIN($E$2:$E$20);$E2=MAX($E$2:$E$20))

Význam:

nejlepší NEBO nejhorší výsledek
Textové podmínky
HLEDAT
=HLEDAT(4;$D2)=1

Vrací:

pozici textu
ZLEVA
=ZLEVA($D2;1)*1=4

Postup:

vezme první znak
převede text na číslo
porovná s hodnotou 4
BI mindset

Podmíněné formátování:

není jen design
je to alerting systém
zvýrazňuje problémy
zrychluje interpretaci dashboardu

🔢 42. ZVOLIT() a INDEX()
ZVOLIT()

Syntaxe:

=ZVOLIT(index;hodnota1;hodnota2;...)

Příklad:

=ZVOLIT(2;"LOW";"MEDIUM";"HIGH")

Výsledek:

MEDIUM
INDEX()

Syntaxe:

=INDEX(oblast;řádek;sloupec)

Příklad:

=INDEX(A1:C10;3;2)

Vrací:

hodnotu ze 3. řádku
a 2. sloupce
Vrácení celého sloupce
=INDEX(B2:D20;0;2)

Význam:

0 v argumentu řádek
vrátí celý 2. sloupec
Vrácení celého řádku
=INDEX(A2:D20;5;0)

Význam:

vrátí celý 5. řádek
INDEX + POZVYHLEDAT = profesionální lookup
=INDEX(B2:B10;POZVYHLEDAT(E2;A2:A10;0))

Výhody:

lookup doleva
stabilní při změně sloupců
flexibilnější než SVYHLEDAT
BI mindset

INDEX():

vybírá data dynamicky
odděluje logiku hledání a vrácení hodnoty
je základ robustních lookupů

🧮 43. SUBTOTAL() a AGGREGATE()
SUBTOTAL()

Syntaxe:

=SUBTOTAL(konstanta;oblast)
Nejčastější konstanty
Konstanta	Význam
1	PRŮMĚR
2	COUNT
4	MAX
5	MIN
9	SUMA
109	SUMA bez skrytých řádků
Součet s ignorováním skrytých řádků
=SUBTOTAL(109;C2:C20)
AGGREGATE()

Syntaxe:

=AGGREGATE(funkce;možnosti;oblast)
Příklad
=AGGREGATE(9;7;C2:C20)

Význam:

9 = SUMA
7 = ignoruj:
skryté řádky
chyby
vnořené agregace
Rozdíl SUMA vs SUBTOTAL
Funkce	Filtr	Skryté řádky	Chyby
SUMA	❌	❌	❌
SUBTOTAL	✅	částečně	❌
AGGREGATE	✅	✅	✅
BI mindset

SUBTOTAL():

reaguje na filtr
ideální pro interaktivní reporting

AGGREGATE():

robustní reporting nad nečistými daty

📑 44. Automatické souhrny dat
Souhrny

Data → Přehled → Souhrn

Excel automaticky:

vloží mezisoučty
vytvoří SUBTOTAL()
vytvoří sbalovací hierarchii
Podmínka správného fungování

Data musí být:

seskupena
seřazena podle kategorií
Celkový součet více oblastí
=SUBTOTAL(9;C2:C13;D2:D13;E2:E13)

Výhoda:

ignoruje mezisoučty uvnitř oblasti
Dynamická změna výpočtu

Dropdown:

Součet
Průměr
Maximum
Minimum

Lookup pozice:

=POZVYHLEDAT(H1;A18:A28;0)

Použití:

dynamický reporting
user-driven dashboard
Ověření dat

Data → Ověření dat → Seznam

Použití:

dropdowny
řízení parametrů reportu
omezení chyb uživatele
BI mindset

Souhrny:

vytváří hierarchii reportu
umožňují drill-down analýzu
oddělují detail a management overview

🔄 45. POSUN() a dynamické oblasti
POSUN()

Syntaxe:

=POSUN(odkaz;řádky;sloupce;výška;šířka)
Jednoduchý posun
=POSUN(A1;2;1)

Výsledek:

vrátí odkaz na B3
Dynamická oblast
=POSUN(C1;0;0;POČET2(C:C);1)

Význam:

oblast se automaticky rozšiřuje
podle počtu hodnot ve sloupci
Dynamický dropdown

Použití:

ověření dat
automatické seznamy
dashboard filtry
Součet posledních X hodnot
=SUMA(POSUN(B1;POČET2(B:B)-E1;0;E1;1))

Význam:

E1 = počet posledních hodnot
dynamický rolling calculation
Volatilní funkce

POSUN() je volatilní:

přepočítává se při každé změně sešitu
u velkých souborů může zpomalovat Excel

Alternativa:

INDEX()
Excel tabulky
dynamická pole
BI mindset

Dynamické oblasti:

automaticky reagují na nová data
minimalizují manuální údržbu reportů
jsou základ automatizovaných dashboardů
🧠 46. Excel Advanced – Celkový mindset
Junior analytik workflow
RAW DATA
↓
Excel Table
↓
Čištění dat
↓
Lookupy
↓
Agregace
↓
Podmíněná logika
↓
Souhrny
↓
Dashboard
↓
Power BI
Nejčastější reálné kombinace
Lookup + IFERROR
=IFERROR(XLOOKUP(A2;Tabulka[ID];Tabulka[Jméno]);"Nenalezeno")
INDEX + MATCH + Dashboard
=INDEX(Data[Tržba];POZVYHLEDAT(H2;Data[Produkt];0))
SUMIFS + datumové podmínky
=SUMIFS(D:D;A:A;">="&H1;A:A;"<="&H2)
KPI alert
=A([@Marže]<0,2;[@Reklamace]>3)

→ použito v podmíněném formátování

🚀 Cíl advanced Excelu

Neumět jen vzorce.

Ale:

automatizovat reporting
minimalizovat ruční práci
vytvářet robustní modely
přemýšlet datově
připravovat data pro Power BI a SQL
