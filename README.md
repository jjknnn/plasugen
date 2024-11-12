# plasugen.py

Luo istumajärjestys tapahtumaasi vaivattomasti! Scipti Toimii mutta bugivapaustakuuta en anna!
Eikun vaan kokeilemaan!

Varmaan jossain vaiheessa luon tästä suoritettavan .exe tiedoston.

##


### Käyttöohjeet

**Kaverisitsien ohjeet alempana**

Ajaaksesi plasugen.py scriptin tarvitset nimilistan, pythonin
ja gitin asennettuna laitteellesi  ja pääsyn komentoriville. 

1. Navigoi komentorivillä plasugen kansioon.
2. Kopioi nimilista kansioon ja nimeä se ```nimilista.xlsx```. Tyypin
tulee olla ```.xlsx``` Nimilistan sisällön tulee olla seuraavanlainen:
Nimikantarivi tulee jättää tyhjäksi, eli xlsx tiedostoon VAIN nimet 

| nimikanta1  | nimikanta2 |nimikanta3 |...
| ----------  |:----------:|:---------:|:-:
| Nimi Niminen|    ...     |    ...
| Nani Naninen|    ...     |    ...
|     ...     |    ...     |    ...

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

Scripti kysyy seuraavia asioita:
1. Random / Järjestys:
    1. Random -> Sekoittaa kaikki nimet ja arpoo paikat
    täysin satunnaisesti
    2. Järjestys -> Luo paikat vuorotellen pöytiin. Jos nimilistat
    eripituiset huomioi pituudet ja muokkaa järjestyksen tasaiseksi
2. Pöytien koot ja määrät:
    1. Koko: Kuinka monta paikkaa pöydässä
    2. Paikkamäärä: Kuinka monta edellä annetun kokoista pöytää
    3. Jatketaanko y/n: Jos haluat luoda lisää pöytiä syötä y, 
    jos kaikki pöydät syötetty, syötä n

### Kaverisitsit

Kaverisitsit toimivat pitkälti samalla tavalla kuin ylempä.
Lue ensiksi se, ja sitten palaa tähän.

Nimilistan muotoilu on hieman erilainen. Muotoile nimet tiedostoon ```nimilista.xlsx``` seuraavanlaisesti:

| Ryhmänjäsen1| jäsen2          |jäsen3        | jäsen4 | ja niin edelleen |
| ----------  |:----------:     |:---------:   |  :-:   |         ---      |
| Mortti      | Mortin kaveri1  |Mortin kaveri2| ...    |                  |
| Nanni       | Nannin kaveiri1 |    ...       | ...    |                  |
|     ...     |    ...          |    ...       | ...    |                  |

Scripti kysyy muuten samat asiat, mutta syötä kysymylseen random/järjestys/kaveri kirjain ```k```

##### Scripti saattaa antaa erroreita. Saa nostaa issueiksi jos haluaa
    
    
    
    
    
    
    
    
    
    
    
