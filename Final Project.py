#Oppgave a)


# NB!!!!!!!!!!!!!! Alt i denne oppgaven avhenger av at filen "support_uke_24.xlsx" Ligger i working directory

"""Del a) Skriv et program som leser inn filen ‘support_uke_24.xlsx’ og lagrer data fra kolonne 1
i en array med variablenavn ‘u_dag’, dataen i kolonne 2 lagres i arrayen ‘kl_slett’, data i
kolonne 3 lagres i arrayen ‘varighet’ og dataen i kolonne 4 lagres i arrayen ‘score’. Merk:
filen ‘support_uke_24.xlsx’ må ligge i samme mappe som Python-programmet ditt."""

import pandas as pd
import numpy as np

# Les Excel-filen
filnavn = 'support_uke_24.xlsx'

# Les inn alle kolonnene med pandas
data = pd.read_excel(filnavn)

# Hent ut kolonnene og lagre i forskjellige arrays
# Ukedag (kolonne 1) – hvilken dag henvendelsen kom inn (Mandag, Tirsdag, ...)
# Bruker .iloc[:, 0] for å hente første kolonne
u_dag = data.iloc[:, 0].to_numpy()

# Klokkeslett (kolonne 2) – tidspunkt på døgnet da kunden tok kontakt
kl_slett = data.iloc[:, 1].to_numpy()

# Samtalens varighet (kolonne 3) – hvor lenge samtalen varte
varighet = data.iloc[:, 2].to_numpy()

# Tilfredshetsscore (kolonne 4) – kundens tilbakemelding, skala 1–10
# Mange rader mangler denne verdien (blir NaN), og det må håndteres senere
score = data.iloc[:, 3].to_numpy()

# Skriver ut for å bekrefte at alt fungerer
print("Ukedag:", u_dag)
print("Klokkeslett:", kl_slett)
print("Varighet:", varighet)
print("Score:", score)
# Det fungerte fint, skriver ut "nan" der det ikke er noen verdi, "Not a Number"


#%%
#Oppgave b)

"""Del b) Skriv et program som finner antall henvendelser for hver de 5 ukedagene. Resultatet
visualiseres ved bruk av et søylediagram (stolpediagram)."""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Leser Excel-filen
filnavn = 'support_uke_24.xlsx'
data = pd.read_excel(filnavn)

# Hent ut ukedager
u_dag = data.iloc[:, 0].to_numpy()

# Tell antall henvendelser per ukedag
ukedag, antall_henvendelser = np.unique(u_dag, return_counts=True)

# Sorterer ukedagene manuelt i riktig rekkefølge
ukedag_rekkefølge = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"]
dag_til_antall = dict(zip(ukedag, antall_henvendelser))
sortert_antall = [dag_til_antall.get(dag, 0) for dag in ukedag_rekkefølge]

# Lager og plotter et stolpediagram
plt.figure(figsize=(8, 5))
plt.bar(ukedag_rekkefølge, sortert_antall, color='skyblue')
plt.title('Antall supporthenvendelser per ukedag')
plt.xlabel('Ukedag')
plt.ylabel('Antall henvendelser')
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()

#%%
#Oppgave c)

"""Del c) Skriv et program som finner minste og lengste samtaletid som er loggført for uke 24.
Svaret skrives til skjerm med informativ tekst."""

import pandas as pd
import numpy as np

# Leser Excel-filen
filnavn = 'support_uke_24.xlsx'
data = pd.read_excel(filnavn)

# Henter ut samtalevarighet som en aray 
varighet = data.iloc[:, 2].to_numpy()

# Finner minimum og maksimum samtaletid
korteste = np.min(varighet)
lengste = np.max(varighet)

# Skriv resultat til skjerm
print(f"Den korteste samtalen var på {korteste} minutter.")
print(f"Den lengste samtalen var på {lengste} minutter.")
# Kjørte denne og dobbelt sjekket ved å sortere kolonnen i Excel, og fikk samme svar. Den fungerer!


#%%
#Oppgave d)

"""Del d) Skriv et program som regner ut gjennomsnittlig samtaletid basert på alle
henvendelser i uke 24."""

import pandas as pd
import numpy as np

# Les Excel-filen
filnavn = 'support_uke_24.xlsx'
data = pd.read_excel(filnavn)

# Konverter 'Varighet'-kolonnen (kolonne 3) til tidsverdier
varighet_tid = pd.to_datetime(data.iloc[:, 2], format='%H:%M:%S', errors='coerce')

# Beregn varighet i minutter som floats
varighet_minutter = varighet_tid.dt.hour * 60 + varighet_tid.dt.minute + varighet_tid.dt.second / 60

# Filtrer bort NaN (tomme rader)
gyldige_varigheter = varighet_minutter.dropna()

# Sjekk om det finnes gyldige data
if gyldige_varigheter.empty:
    print("Ingen gyldige samtaletider funnet.")
else:
    gjennomsnitt = gyldige_varigheter.mean()
    print(f"Gjennomsnittlig samtaletid i uke 24 var {gjennomsnitt:.2f} minutter.")

# Ja, denne var litt knot å få til å skrive ut korrekt.  Jeg fant ingen måte å dobbeltsjekke denne, men med tanke på varighetene ser det riktig ut. 
    
    
#%%
#Oppgave e)

"""Del e) Supportvaktene i MORSE er delt inn i 2-timers bolker: kl 08-10, kl 10-12, kl 12-14 og kl
14-16. Skriv et program som finner det totale antall henvendelser supportavdelingen mottok
for hver av tidsrommene 08-10, 10-12, 12-14 og 14-16 for uke 24. Resultatet visualiseres ved
bruk av et sektordiagram (kakediagram)."""

import pandas as pd
import matplotlib.pyplot as plt

# Les Excel-filen
filnavn = 'support_uke_24.xlsx'
data = pd.read_excel(filnavn)

# Konverter klokkeslett (kolonne 2) til datetime
tidspunkt = pd.to_datetime(data.iloc[:, 1], format='%H:%M:%S', errors='coerce')

# Ekstraher timer
timer = tidspunkt.dt.hour

# TELLE HENVENDELSER I ULIKE TIDSSONER (2-TIMERS INTERVALLER)

# Her bruker vi dictionary for å gruppere antall henvendelser
# Variabelen 'timer' inneholder timeverdien (klokkeslett) fra hver henvendelse

intervaller = {
    # 08:00–09:59 → inkluderer alle henvendelser fra og med kl 08:00 til før 10:00
    "08–10": ((timer >= 8) & (timer < 10)).sum(),

    # 10:00–11:59 → inkluderer alle henvendelser fra og med kl 10:00 til før 12:00
    "10–12": ((timer >= 10) & (timer < 12)).sum(),

    # 12:00–13:59 → inkluderer alle henvendelser fra og med kl 12:00 til før 14:00
    "12–14": ((timer >= 12) & (timer < 14)).sum(),

    # 14:00–15:59 → inkluderer alle henvendelser fra og med kl 14:00 til før 16:00
    "14–16": ((timer >= 14) & (timer < 16)).sum(),
}

# Hver linje filtrerer time-verdiene og teller antall True-verdier (dvs. treff),
# ved hjelp av sum() på en boolsk maske.


# Totalen
total_antall = sum(intervaller.values())

# Funksjon for å vise både prosent og antall
def vis_prosent_og_antall(pct, all_vals):
    absolute = int(round(pct / 100. * sum(all_vals)))
    return f"{pct:.1f}%\n({absolute})"

# Tegn sektordiagram
plt.figure(figsize=(6, 6))
plt.pie(
    intervaller.values(),
    labels=intervaller.keys(),
    autopct=lambda pct: vis_prosent_og_antall(pct, list(intervaller.values())),
    startangle=90
)
plt.title('Supporthenvendelser per tidsrom\n(Prosent og antall)')
plt.axis('equal')  # Gjør sirkelen rund

# Legg til tekstboks med totalen
plt.text(-1.4, -1.1, f"Sum antall henvendelser: {total_antall}", fontsize=10, ha='left')

plt.tight_layout()
plt.show()




#%%
#Oppgave f)

"""Del f) Kundens tilfredshet loggføres som tall fra 1-10 hvor 1 indikerer svært misfornøyd og
10 indikerer svært fornøyd. Disse tilbakemeldingene skal så overføres til NPS-systemet (Net
Promoter Score)."""


import pandas as pd
import numpy as np

# Les Excel-filen
filnavn = 'support_uke_24.xlsx'
data = pd.read_excel(filnavn)

# Hent tilfredshet (kolonne 4) og filtrer bort tomme rader
score_raw = data.iloc[:, 3]
score = pd.to_numeric(score_raw, errors='coerce').dropna()

# Tell kategorier
antall_total = len(score)
antall_positive = ((score >= 9) & (score <= 10)).sum()
antall_nøytrale = ((score >= 7) & (score <= 8)).sum()
antall_negative = ((score >= 1) & (score <= 6)).sum()

# Beregn prosentandeler
prosent_positive = (antall_positive / antall_total) * 100
prosent_negative = (antall_negative / antall_total) * 100

# Beregn NPS
nps = prosent_positive - prosent_negative

# Skriv resultater til skjerm
print(f"Totalt {antall_total} kuner har gitt tilbakemelding.")
print(f"Positive (9–10): {antall_positive} ({prosent_positive:.1f}%)")
print(f"Nøytrale (7–8): {antall_nøytrale}")
print(f"Negative (1–6): {antall_negative} ({prosent_negative:.1f}%)")
print(f"\n📈 Supportavdelingens nPS = {nps:.1f}")

# Litt usikker på om jeg har forstått utregningen her, men siden det kun er NPS = positive - negative som da blir 
# NPS = 65.1 − 23.3 = 41.8 skjønner jeg ikke helt hvordan det det kan bli anderledes""" 



