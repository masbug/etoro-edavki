
# eToro -> FURS eDavki konverter

_Konverter iz eToroAccountStatement-a v XLSX (MS Excel) obliki pripravi datoteke v XML format primerne za uvoz v eDavke:_
* _Doh-KDVP - Napoved za odmero dohodnine od dobička od odsvojitve vrednostnih papirjev in drugih deležev ter investicijskih kuponov,_
* _D-IFI - Napoved za odmero davka od dobička od odsvojitve izvedenih finančnih instrumentov_
* _Doh-Div - Napoved za odmero dohodnine od dividend_

Skripta avtomatsko naredi konverzijo tuje valute v EUR po tečaju Banke Slovenije na dan posla.

## Izjava o omejitvi odgovornosti

Davki so resna stvar. Avtor(ji) skripte si prizadevam(o) za natančno in ažurno delovanje skripte in jo tudi sam(i)
uporabljam(o) za napovedi davkov. Kljub temu ne izključujem(o) možnosti napak, ki lahko vodijo v napačno oddajo davčne
napovedi. Za pravilnost davčne napovedi si odgovoren sam in avtor(ji) skripte ne prevzema(mo) nobene odgovornosti.

Pisanje skripte mi je vzelo precej časa. Sam donacij ne sprejemam, bom pa vesel vsake donacije, ki jo podelite kateri od dobrodelnih organizacij, npr. [Slovenski Karitas](https://www.karitas.si/daruj/donacija/).

## Uporaba

### Namestitev skripte

```
pip3 install --upgrade git+https://github.com/masbug/etoro-edavki.git
```

```
etoro-edavki
```

ali pa si prenesete že prevedeno skripto (za uporabo glej PREBERI_ME.txt):
https://github.com/masbug/etoro-edavki/releases

### Izvoz poročila na eToro

1. V meniju odpri **Portfolio**
2. Desno od menija poleg besede **Portfolio** klikni na "urico" (_history_).
3. V history pogledu klikni na zobnik skrajno desno in izberi "Account statement".
4. Vpiši začetni datum (01/01/_prejšnje leto, bolje še kakšno prej_) in končni datum (01/01/_to leto_).
5. Klikni na kljukico za potrditev.
6. Klikni na XLS ikono za izvoz v Excel obliki.

### Konverzija poročila v popisne liste primerne za uvoz v eDavke

```
etoro-edavki [-h] [-c] [-y report-year] eToroAccountStatement-2024.xlsx
```
Argumenti:
*    -y: ročno izbere leto za katero se naj XMLji izvozijo (debugging)
*    -c: vključi tudi "real" kripto pozicije v napovedi (CFD so vedno vključene)
*    eToroAccountStatement-2024.xlsx: datoteka, ki jo prenesemo iz eToro

#### Postopek
Skripta najprej avtomatsko prenese tabelo za konverzijo valut, nato v mapi output ustvari 4 datoteke:
* **Doh-KDVP.xml** (datoteka namenjena uvozu v obrazec **Doh-KDVP** - Napoved za odmero dohodnine od dobička od odsvojitve vrednostnih papirjev in drugih deležev ter investicijskih kuponov)
* **D-IFI.xml** (datoteka namenjena uvozu v obrazec **D-IFI** - Napoved za odmero davka od dobička od odsvojitve izvedenih finančnih instrumentov)
* **Doh-Div.xml** (datoteka namenjena uvozu v obrazec **Doh-Div**)

* **Dividende-info-**_leto_**.xlsx** (kontrolna datoteka; v pomoč pri hitrem pregledu manjkajočih podatkov za generiranje Doh-Div)
* Debug-_leto_.xlsx (kontrolna datoteka za Doh-KDVP, D-IFI)

#### Obrazec Doh-Div
Obrazec Doh-Div zahteva dodatne podatke o podjetju, ki je izplačalo dividende (identifikacijska številka, naslov, ISIN), ki jih v izvirnih podatkih eTora ni. Te podatke je potrebno ročno poiskati in dopisati v Naslovi_info.xlsx.

~~Pri uveljavljanju olajšave za že odvedeni davek v tujini, je potrebno na eDavkih specificirati mednarodno pogodbo od preprečevanju dvojnega obdavčevanja in člen. Za pomoč pri pogodbah je tabela:~~
~~- https://www.gov.si/drzavni-organi/ministrstva/ministrstvo-za-finance/o-ministrstvu/direktorat-za-sistem-davcnih-carinskih-in-drugih-javnih-prihodkov/seznam-veljavnih-konvencij-o-izogibanju-dvojnega-obdavcevanja-dohodka-in-premozenja/~~
~~...člene pa je potrebno ročno poiskat.~~

[Po zadnjih informacijah ni potrebno izpolniti polja za mednarodno pogodbo.](https://github.com/masbug/etoro-edavki/issues/28#issuecomment-1899933818)

### Uvoz v eDavke
1. V meniju **Dokument** klikni **Uvoz**. Izberi eno izmed generiranih datotek (Doh-KDVP.xml, D-IFI, Doh-Div) in jo **Prenesi**.
![image](https://user-images.githubusercontent.com/11191264/221360416-bce47565-7f11-4d5e-a3e2-4466785c1bc7.png)
2. Preveri izpolnjene podatke in dodaj manjkajoče.
3. Pri obrazcih Doh-KDVP in D-IFI je na seznamu popisnih listov po en popisni list za vsak vrednostni papir (ticker).
4. Klikni na ime vrednostnega papirja in odpri popisni list.
5. Klikni **Izračun**.
6. Preveri če vse pridobitve in odsvojitve ustrezajo dejanskim. Zaloga pri zadnjem vnosu mora biti **0**.

### Viri
_Osnova za to skripto je IB -> eDavki konverter: https://github.com/jamsix/ib-edavki (Copyright (c) 2020 Primož Sečnik Kolman; MIT Licenca)_
