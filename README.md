# plasugen.py

Luo istumajärjestys tapahtumaasi vaivattomasti! Scipti toimii mutta bugivapauteen takuuta en anna!
Eikun vaan kokeilemaan!

Suoritettava .exe tiedoso on nyt luotu! Käytä plasugenia suoraan suoritettavasta tiedostosta
ilman ylimääräistä säätöä!


## .exe eli yksinkertainen käyttö

Tarvitset vain Windows käyttöjärjestelmän!

1. Lataa ja suorita plasugen.exe
2. Valitse käytettävä .xlsx tiedosto. 
3. Scripti kysyy seuraavia asioita:
    1. Random / Järjestys / Kaveri:
        - Random -> Sekoittaa kaikki nimet ja arpoo paikat
        täysin satunnaisesti
        - Järjestys -> Luo paikat vuorotellen eri nimikannoista 
        (esim ainejärjestöt) tasaisin välein pöytiin. Jos
        nimilistat eripituiset huomioi pituudet ja muokkaa
        järjestyksen tasaiseksi.
        - Kaveri -> Kaverisitsiformaatti. Jakaa kaveriporukat
        vierekkäin pöytiin. **HUOM!** Suuret nimilistat voivat viedä
        paljon aikaa!
    2. Pöytien koot ja määrät:
        - Koko: Kuinka monta paikkaa pöydässä
        - Paikkamäärä: Kuinka monta edellä annetun kokoista pöytää
        - Jatketaanko y/n: Jos haluat luoda lisää pöytiä syötä y, 
        jos kaikki pöydät syötetty, syötä n
4. Scipti generoi ```names.xlsx``` nimisen tiedoston siihen kansioon,
    josta .exe löytyy!


## Nimilistojen muotoilu

### Muut kuin kaverisitsit

Tyypin tulee olla ```.xlsx``` Nimilistan sisällön tulee olla
seuraavanlainen:

| nimikanta1  | nimikanta2 |nimikanta3 |...
| ----------  |:----------:|:---------:|:-:
| Nimi Niminen|    ...     |    ...
| Nani Naninen|    ...     |    ...
|     ...     |    ...     |    ...

Nimikantarivi, eli osallistujat jaettuna viiteryhmittäin 
(esim ainejärjestöt) tulee jättää tyhjäksi, eli xlsx tiedostoon VAIN nimet

### Kaverisitsit

Kaverisitsit toimivat pitkälti samalla tavalla kuin ylempä.
Lue ensiksi se, ja sitten palaa tähän.

Nimilistan muotoilu on hieman erilainen. Muotoile nimet tiedostoon ```nimilista.xlsx``` seuraavanlaisesti:

| Ryhmänjäsen1| jäsen2          |jäsen3        | jäsen4 | ja niin edelleen |
| ----------  |:----------:     |:---------:   |  :-:   |         ---      |
| Mortti      | Mortin kaveri1  |Mortin kaveri2| ...    |                  |
| Nanni       | Nannin kaveiri1 |    ...       | ...    |                  |
|     ...     |    ...          |    ...       | ...    |                  |

Scripti kysyy muuten samat asiat, mutta syötä kysymykseen random/järjestys/kaveri kirjain ```k```.

Ohjelman suoritusaika riippuu suoraan nimilistan koosta ja kasvaa eksponentiaalisesti.
Isot, yli 120 hlö kaverisitsit voivat viedä paljon aikaa.

Mikäli tuntuu, että ohjelma ei suoritu minuuttien jälkeen kannattaa se keskeyttää, luoda itse muutama
pöytä valmiiksi ja kokeilla uudestaan.

## Scriptin ajaminen pythonilla

Ajaaksesi plasugen.py scriptin tarvitset nimilistan, pythonin
ja gitin asennettuna laitteellesi sekä pääsyn komentoriville. 

1. Navigoi komentorivillä plasugen kansioon.
2. Kopioi nimilista kansioon ja nimeä se ```nimilista.xlsx```. 
3. Navigoi komentorivillä kansioon missä haluat tiedostojen sijaitsevan
ja kopioi tiedostot
    ```
    git clone https://github.com/jjknnn/plasugen.git
    ```


4. Asennetaan tarvittavat paketit pythoniin:
    ```
    python3 -m pip install -r requirements.txt
    ```

5. Aja scipti:

    ```
    python3 plasugen.py
    ```
6. Scipti generoi ```names.xlsx``` nimisen tiedoston, josta löytyy 
pöydät nimineen


##### Scripti saattaa antaa erroreita. Saa nostaa issueiksi jos haluaa
    
