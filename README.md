# plasugen.py

Luo istumajärjestys tapahtumaasi vaivattomasti! Scipti on vielä
kehityksessä! Innokkaat istumajärjestyksen luojat voivat toki
jo kokeilla!

##


#### Käyttöohjeet

Ajaaksesi plasugen.py scriptin tarvitset nimilistan, pythonin
asennettuna laitteellesi  ja pääsyn komentoriville. 

1. Navigoi komentorivillä plasugen kansioon.
2. Kopioi nimilista kansioon ja nimeä se ```nimilista.xlsx```. Tyypin
tulee olla ```.xlsx``` Nimilistan sisällön tulee olla seuraavanlainen:
Nimikantarivi tulee jättää tyhjäksi, eli xlsx tiedostoon VAIN nimet 

| nimikanta1  | nimikanta2 |nimikanta3 |...
| ----------  |:----------:|:---------:|:-:
| Nimi Niminen|    ...     |    ...
| Nani Naninen|    ...     |    ...
|     ...     |    ...     |    ...
3. Asennetaan tarvittavat paketit pythoniin:
    ```
    python3 -m pip install requirements.txt
    ```

4. Aja scipti:

    ```
    python3 plasugen.py
    ```
5. Scipti generoi ```names.xlsx``` nimisen tiedoston, josta löytyy 
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

##### Scripti saattaa antaa erroreita ja korjailen bugeja kunhan kerkiän
    
    
    
    
    
    
    
    
    
    
    
