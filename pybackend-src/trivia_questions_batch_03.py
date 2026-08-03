# -*- coding: utf-8 -*-
# Trivia Challenge batch #3. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which orange root vegetable is often roasted or mashed as a side dish?', 'de': 'Welches orangefarbene Wurzelgemuse wird oft als Beilage geroestet oder pueriert?', 'es': '¿Que verdura de raiz naranja se asa o hace pure como guarnicion?', 'nl': 'Welke oranje wortelgroente wordt vaak geroosterd of gepureerd als bijgerecht?', 'sv': 'Vilken orange rotfrukt rostas eller mosas ofta som tillbehor?', 'fr': "Quel legume-racine orange est souvent roti ou reduit en puree comme accompagnement ?", 'da': 'Hvilken orange rodfrugt ristes eller mostes ofte som tilbehor?'},
 'opt1': {'en': 'Carrot', 'de': 'Karotte', 'es': 'Zanahoria', 'nl': 'Wortel', 'sv': 'Morot', 'fr': 'Carotte', 'da': 'Gulerod'},
 'opt2': {'en': 'Parsnip', 'de': 'Pastinake', 'es': 'Chirivia', 'nl': 'Pastinaak', 'sv': 'Palsternacka', 'fr': 'Panais', 'da': 'Pastinak'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is famous for the Colosseum?', 'de': 'Welches Land ist beruehmt fuer das Kolosseum?', 'es': '¿Que pais es famoso por el Coliseo?', 'nl': 'Welk land is beroemd om het Colosseum?', 'sv': 'Vilket land ar kant for Colosseum?', 'fr': "Quel pays est celebre pour le Colisee ?", 'da': 'Hvilket land er beromt for Colosseum?'},
 'opt1': {'en': 'Italy', 'de': 'Italien', 'es': 'Italia', 'nl': 'Italie', 'sv': 'Italien', 'fr': "l'Italie", 'da': 'Italien'},
 'opt2': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a scarecrow, a tin man, and a cowardly lion?', 'de': 'In welchem Filmklassiker kommen eine Vogelscheuche, ein Blechmann und ein feiger Loewe vor?', 'es': '¿Que pelicula clasica presenta a un espantapajaros, un hombre de hojalata y un leon cobarde?', 'nl': 'In welke klassieke film komen een vogelverschrikker, een blikken man en een lafhartige leeuw voor?', 'sv': 'I vilken klassisk film forekommer en fagelskrammel, en plaman och ett feigt lejon?', 'fr': "Quel film classique met en scene un epouvantail, un homme de fer-blanc et un lion peureux ?", 'da': 'Hvilken klassisk film har en fugleskræmsel, en blikmand og en frygtsom løve?'},
 'opt1': {'en': 'The Wizard of Oz', 'de': 'Der Zauberer von Oz', 'es': 'El Mago de Oz', 'nl': 'The Wizard of Oz', 'sv': 'Trollkarlen fran Oz', 'fr': "Le Magicien d'Oz", 'da': 'Troldmanden fra Oz'},
 'opt2': {'en': 'Peter Pan', 'de': 'Peter Pan', 'es': 'Peter Pan', 'nl': 'Peter Pan', 'sv': 'Peter Pan', 'fr': 'Peter Pan', 'da': 'Peter Pan'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which pet is known for barking and is often called 'man's best friend'?", 'de': "Welches Haustier bellt und wird oft 'des Menschen bester Freund' genannt?", 'es': "¿Que mascota ladra y a menudo se le llama 'el mejor amigo del hombre'?", 'nl': "Welk huisdier blaft en wordt vaak 'de beste vriend van de mens' genoemd?", 'sv': "Vilket sallskapsdjur skaller och kallas ofta 'manniskans basta van'?", 'fr': "Quel animal de compagnie aboie et est souvent appele le 'meilleur ami de l'homme' ?", 'da': "Hvilket kæledyr gør og kaldes ofte 'menneskets bedste ven'?"},
 'opt1': {'en': 'Dog', 'de': 'Hund', 'es': 'Perro', 'nl': 'Hond', 'sv': 'Hund', 'fr': 'Chien', 'da': 'Hund'},
 'opt2': {'en': 'Cat', 'de': 'Katze', 'es': 'Gato', 'nl': 'Kat', 'sv': 'Katt', 'fr': 'Chat', 'da': 'Kat'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which baked good is traditionally topped with frosting and candles for birthdays?', 'de': 'Welches Gebaeck wird traditionell mit Zuckerguss und Kerzen zum Geburtstag verziert?', 'es': '¿Que producto horneado se cubre tradicionalmente con glaseado y velas en cumpleaños?', 'nl': 'Welk gebak wordt traditioneel bedekt met glazuur en kaarsjes voor verjaardagen?', 'sv': 'Vilket bakverk toppas traditionellt med glasyr och ljus pa fodelsedagar?', 'fr': "Quelle patisserie est traditionnellement recouverte de glacage et de bougies pour les anniversaires ?", 'da': 'Hvilket bagvaerk toppes traditionelt med glasur og lys til fodselsdage?'},
 'opt1': {'en': 'Cake', 'de': 'Kuchen', 'es': 'Pastel', 'nl': 'Taart', 'sv': 'Tarta', 'fr': 'Gateau', 'da': 'Kage'},
 'opt2': {'en': 'Pie', 'de': 'Torte', 'es': 'Tarta', 'nl': 'Pastei', 'sv': 'Pai', 'fr': 'Tarte', 'da': 'Tærte'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to giant pandas in the wild?', 'de': 'In welchem Land leben Riesenpandas in freier Wildbahn?', 'es': '¿En que pais viven pandas gigantes en estado salvaje?', 'nl': 'In welk land leven reuzenpanda\'s in het wild?', 'sv': 'I vilket land lever jattepandor i det vilda?', 'fr': "Dans quel pays vivent les pandas geants a l'etat sauvage ?", 'da': 'Hvilket land huser kæmpepandaer i naturen?'},
 'opt1': {'en': 'China', 'de': 'China', 'es': 'China', 'nl': 'China', 'sv': 'Kina', 'fr': 'la Chine', 'da': 'Kina'},
 'opt2': {'en': 'Mongolia', 'de': 'Mongolei', 'es': 'Mongolia', 'nl': 'Mongolie', 'sv': 'Mongoliet', 'fr': 'la Mongolie', 'da': 'Mongoliet'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. president is depicted on Mount Rushmore alongside Washington, Jefferson, and Lincoln?', 'de': 'Welcher US-Praesident ist am Mount Rushmore neben Washington, Jefferson und Lincoln abgebildet?', 'es': '¿Que presidente de EE.UU. esta representado en el Monte Rushmore junto a Washington, Jefferson y Lincoln?', 'nl': 'Welke Amerikaanse president is afgebeeld op Mount Rushmore naast Washington, Jefferson en Lincoln?', 'sv': 'Vilken amerikansk president avbildas pa Mount Rushmore tillsammans med Washington, Jefferson och Lincoln?', 'fr': "Quel president americain figure sur le mont Rushmore aux cotes de Washington, Jefferson et Lincoln ?", 'da': 'Hvilken amerikansk præsident er afbildet på Mount Rushmore sammen med Washington, Jefferson og Lincoln?'},
 'opt1': {'en': 'Theodore Roosevelt', 'de': 'Theodore Roosevelt', 'es': 'Theodore Roosevelt', 'nl': 'Theodore Roosevelt', 'sv': 'Theodore Roosevelt', 'fr': 'Theodore Roosevelt', 'da': 'Theodore Roosevelt'},
 'opt2': {'en': 'Woodrow Wilson', 'de': 'Woodrow Wilson', 'es': 'Woodrow Wilson', 'nl': 'Woodrow Wilson', 'sv': 'Woodrow Wilson', 'fr': 'Woodrow Wilson', 'da': 'Woodrow Wilson'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which art style is known for melting clocks, associated with Salvador Dali?', 'de': 'Welcher Kunststil ist fuer schmelzende Uhren bekannt und mit Salvador Dali assoziiert?', 'es': '¿Que estilo artistico es conocido por los relojes derretidos, asociado con Salvador Dali?', 'nl': 'Welke kunststijl staat bekend om smeltende klokken, geassocieerd met Salvador Dali?', 'sv': 'Vilken konststil ar kand for smaltande klockor, forknippad med Salvador Dali?', 'fr': "Quel style artistique est connu pour ses horloges fondantes, associe a Salvador Dali ?", 'da': 'Hvilken kunststil er kendt for smeltende ure, forbundet med Salvador Dali?'},
 'opt1': {'en': 'Surrealism', 'de': 'Surrealismus', 'es': 'Surrealismo', 'nl': 'Surrealisme', 'sv': 'Surrealism', 'fr': 'Surrealisme', 'da': 'Surrealisme'},
 'opt2': {'en': 'Cubism', 'de': 'Kubismus', 'es': 'Cubismo', 'nl': 'Kubisme', 'sv': 'Kubism', 'fr': 'Cubisme', 'da': 'Kubisme'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which keyboard instrument uses foot pedals and pipes, often found in churches?', 'de': 'Welches Tasteninstrument verwendet Fusspedale und Pfeifen und ist oft in Kirchen zu finden?', 'es': '¿Que instrumento de teclado usa pedales y tubos, y se encuentra a menudo en iglesias?', 'nl': 'Welk toetsinstrument gebruikt voetpedalen en pijpen en is vaak te vinden in kerken?', 'sv': 'Vilket tangentinstrument anvander fotpedaler och pipor och finns ofta i kyrkor?', 'fr': "Quel instrument a clavier utilise des pedales et des tuyaux, souvent present dans les eglises ?", 'da': 'Hvilket tangentinstrument bruger fodpedaler og piber og findes ofte i kirker?'},
 'opt1': {'en': 'Organ', 'de': 'Orgel', 'es': 'Organo', 'nl': 'Orgel', 'sv': 'Orgel', 'fr': 'Orgue', 'da': 'Orgel'},
 'opt2': {'en': 'Harpsichord', 'de': 'Cembalo', 'es': 'Clavecin', 'nl': 'Klavecimbel', 'sv': 'Cembalo', 'fr': 'Clavecin', 'da': 'Cembalo'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which weather event involves a rotating column of air touching the ground?', 'de': 'Welches Wetterereignis besteht aus einer rotierenden Luftsaeule, die den Boden beruehrt?', 'es': '¿Que fenomeno meteorologico consiste en una columna de aire giratoria que toca el suelo?', 'nl': 'Welk weersverschijnsel betreft een draaiende luchtkolom die de grond raakt?', 'sv': 'Vilken vaderhandelse innebar en roterande luftpelare som ror vid marken?', 'fr': "Quel phenomene meteorologique implique une colonne d'air tournante touchant le sol ?", 'da': 'Hvilken vejrbegivenhed involverer en roterende luftsojle, der rorer jorden?'},
 'opt1': {'en': 'Tornado', 'de': 'Tornado', 'es': 'Tornado', 'nl': 'Tornado', 'sv': 'Tornado', 'fr': 'Tornade', 'da': 'Tornado'},
 'opt2': {'en': 'Hurricane', 'de': 'Hurrikan', 'es': 'Huracan', 'nl': 'Orkaan', 'sv': 'Orkan', 'fr': 'Ouragan', 'da': 'Orkan'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic sitcom features a bar owner named Sam Malone?', 'de': 'In welcher klassischen Sitcom kommt ein Barbesitzer namens Sam Malone vor?', 'es': '¿Que sitcom clasica presenta a un dueño de bar llamado Sam Malone?', 'nl': 'In welke klassieke sitcom komt een barhouder genaamd Sam Malone voor?', 'sv': 'I vilken klassisk sitcom forekommer en barägare vid namn Sam Malone?', 'fr': "Quelle sitcom classique met en scene un patron de bar nomme Sam Malone ?", 'da': 'Hvilken klassisk sitcom har en bar-ejer ved navn Sam Malone?'},
 'opt1': {'en': 'Cheers', 'de': 'Cheers', 'es': 'Cheers', 'nl': 'Cheers', 'sv': 'Cheers', 'fr': 'Cheers', 'da': 'Cheers'},
 'opt2': {'en': 'Frasier', 'de': 'Frasier', 'es': 'Frasier', 'nl': 'Frasier', 'sv': 'Frasier', 'fr': 'Frasier', 'da': 'Frasier'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is shaped like a mitten?', 'de': 'Welcher US-Bundesstaat hat die Form eines Fausthandschuhs?', 'es': '¿Que estado de EE.UU. tiene forma de manopla?', 'nl': 'Welke Amerikaanse staat heeft de vorm van een want?', 'sv': 'Vilken amerikansk delstat har formen av en vante?', 'fr': "Quel État americain a la forme d'une moufle ?", 'da': 'Hvilken amerikansk delstat har form som en vante?'},
 'opt1': {'en': 'Michigan', 'de': 'Michigan', 'es': 'Michigan', 'nl': 'Michigan', 'sv': 'Michigan', 'fr': 'le Michigan', 'da': 'Michigan'},
 'opt2': {'en': 'Wisconsin', 'de': 'Wisconsin', 'es': 'Wisconsin', 'nl': 'Wisconsin', 'sv': 'Wisconsin', 'fr': 'le Wisconsin', 'da': 'Wisconsin'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which red fruit is commonly used to make ketchup?', 'de': 'Welche rote Frucht wird ueblicherweise fuer Ketchup verwendet?', 'es': '¿Que fruta roja se usa comunmente para hacer ketchup?', 'nl': 'Welk rood fruit wordt vaak gebruikt om ketchup te maken?', 'sv': 'Vilken rod frukt anvands vanligen for att gora ketchup?', 'fr': "Quel fruit rouge est couramment utilise pour faire du ketchup ?", 'da': 'Hvilken rod frugt bruges almindeligvis til at lave ketchup?'},
 'opt1': {'en': 'Tomato', 'de': 'Tomate', 'es': 'Tomate', 'nl': 'Tomaat', 'sv': 'Tomat', 'fr': 'Tomate', 'da': 'Tomat'},
 'opt2': {'en': 'Bell pepper', 'de': 'Paprika', 'es': 'Pimiento', 'nl': 'Paprika', 'sv': 'Paprika', 'fr': 'Poivron', 'da': 'Peberfrugt'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves two players or teams hitting a ball over a net using rackets, played on a court with a service box?', 'de': 'Bei welcher Sportart schlagen zwei Spieler oder Teams einen Ball mit Schlagern ueber ein Netz, auf einem Platz mit Aufschlagfeld?', 'es': '¿Que deporte consiste en que dos jugadores o equipos golpean una pelota sobre una red con raquetas, en una cancha con zona de saque?', 'nl': 'Bij welke sport slaan twee spelers of teams een bal met rackets over een net, op een baan met een servicevak?', 'sv': 'Vilken sport innebar att tva spelare eller lag slar en boll over ett nat med racketar, pa en bana med en serveruta?', 'fr': "Quel sport implique deux joueurs ou equipes frappant une balle par-dessus un filet avec des raquettes, sur un court avec une zone de service ?", 'da': 'Hvilken sport involverer to spillere eller hold, der slar en bold over et net med ketchere, pa en bane med en servicebox?'},
 'opt1': {'en': 'Tennis', 'de': 'Tennis', 'es': 'Tenis', 'nl': 'Tennis', 'sv': 'Tennis', 'fr': 'Tennis', 'da': 'Tennis'},
 'opt2': {'en': 'Squash', 'de': 'Squash', 'es': 'Squash', 'nl': 'Squash', 'sv': 'Squash', 'fr': 'Squash', 'da': 'Squash'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which insect is known for buzzing and stinging, and lives in a hive?', 'de': 'Welches Insekt summt, sticht und lebt in einem Bienenstock?', 'es': '¿Que insecto zumba, pica y vive en una colmena?', 'nl': 'Welk insect zoemt, steekt en leeft in een korf?', 'sv': 'Vilken insekt surrar, sticks och lever i en kupa?', 'fr': "Quel insecte bourdonne, pique et vit dans une ruche ?", 'da': 'Hvilket insekt summer, stikker og lever i en bikube?'},
 'opt1': {'en': 'Bee', 'de': 'Biene', 'es': 'Abeja', 'nl': 'Bij', 'sv': 'Bi', 'fr': 'Abeille', 'da': 'Bi'},
 'opt2': {'en': 'Fly', 'de': 'Fliege', 'es': 'Mosca', 'nl': 'Vlieg', 'sv': 'Fluga', 'fr': 'Mouche', 'da': 'Flue'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which holiday involves giving thanks and eating a large feast in November in the US?', 'de': 'Bei welchem Feiertag im November in den USA dankt man und isst ein grosses Festmahl?', 'es': '¿Que festividad implica dar gracias y comer un gran banquete en noviembre en EE.UU.?', 'nl': 'Bij welke feestdag in november in de VS geeft men dank en eet men een groot feestmaal?', 'sv': 'Vilken hogtid i november i USA innebar att man tackar och ater en stor fest?', 'fr': "Quelle fete implique de rendre grace et de faire un grand festin en novembre aux États-Unis ?", 'da': 'Hvilken helligdag i november i USA involverer at give tak og spise et stort festmaltid?'},
 'opt1': {'en': 'Thanksgiving', 'de': 'Erntedankfest', 'es': 'Accion de Gracias', 'nl': 'Thanksgiving', 'sv': 'Thanksgiving', 'fr': 'Thanksgiving', 'da': 'Thanksgiving'},
 'opt2': {'en': 'Christmas', 'de': 'Weihnachten', 'es': 'Navidad', 'nl': 'Kerstmis', 'sv': 'Jul', 'fr': 'Noel', 'da': 'Jul'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which ocean is the largest in the world?', 'de': 'Welcher Ozean ist der groesste der Welt?', 'es': '¿Que oceano es el mas grande del mundo?', 'nl': 'Welke oceaan is de grootste ter wereld?', 'sv': 'Vilket hav ar storst i varlden?', 'fr': "Quel ocean est le plus grand du monde ?", 'da': 'Hvilket hav er det storste i verden?'},
 'opt1': {'en': 'Pacific Ocean', 'de': 'Pazifischer Ozean', 'es': 'Oceano Pacifico', 'nl': 'Stille Oceaan', 'sv': 'Stilla havet', 'fr': 'Ocean Pacifique', 'da': 'Stillehavet'},
 'opt2': {'en': 'Indian Ocean', 'de': 'Indischer Ozean', 'es': 'Oceano Indico', 'nl': 'Indische Oceaan', 'sv': 'Indiska oceanen', 'fr': 'Ocean Indien', 'da': 'Det Indiske Ocean'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which author wrote 'Charlotte's Web'?", 'de': "Welcher Autor schrieb 'Charlottes Netz'?", 'es': "¿Que autor escribio 'La Telaraña de Carlota'?", 'nl': "Welke auteur schreef 'Charlotte's Web'?", 'sv': "Vilken forfattare skrev 'Charlottes nat'?", 'fr': "Quel auteur a ecrit 'Le Petit Monde de Charlotte' ?", 'da': "Hvilken forfatter skrev 'Charlottes Web'?"},
 'opt1': {'en': 'E.B. White', 'de': 'E.B. White', 'es': 'E.B. White', 'nl': 'E.B. White', 'sv': 'E.B. White', 'fr': 'E.B. White', 'da': 'E.B. White'},
 'opt2': {'en': 'Beverly Cleary', 'de': 'Beverly Cleary', 'es': 'Beverly Cleary', 'nl': 'Beverly Cleary', 'sv': 'Beverly Cleary', 'fr': 'Beverly Cleary', 'da': 'Beverly Cleary'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which Pixar film features a fish named Nemo?', 'de': 'In welchem Pixar-Film kommt ein Fisch namens Nemo vor?', 'es': '¿Que pelicula de Pixar presenta a un pez llamado Nemo?', 'nl': 'In welke Pixar-film komt een vis genaamd Nemo voor?', 'sv': 'I vilken Pixar-film forekommer en fisk vid namn Nemo?', 'fr': "Quel film Pixar met en scene un poisson nomme Nemo ?", 'da': 'I hvilken Pixar-film optræder en fisk ved navn Nemo?'},
 'opt1': {'en': 'Finding Nemo', 'de': 'Findet Nemo', 'es': 'Buscando a Nemo', 'nl': 'Finding Nemo', 'sv': 'Hitta Nemo', 'fr': 'Le Monde de Nemo', 'da': 'Find Nemo'},
 'opt2': {'en': 'Shark Tale', 'de': 'Große Haie - Kleine Fische', 'es': 'El Espantatiburones', 'nl': 'Shark Tale', 'sv': 'Hajlif', 'fr': 'Gang de requins', 'da': 'Hajen der ville være stjerne'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which breakfast food is made from batter cooked on a griddle and served with syrup?', 'de': 'Welches Fruehstuecksgericht wird aus Teig auf einer Grillplatte gebacken und mit Sirup serviert?', 'es': '¿Que desayuno se hace con masa cocida en una plancha y se sirve con jarabe?', 'nl': 'Welk ontbijtgerecht wordt gemaakt van beslag gebakken op een bakplaat en geserveerd met stroop?', 'sv': 'Vilken frukostmat gors av smet tillagad pa en stekhall och serveras med sirap?', 'fr': "Quel plat de petit-dejeuner est fait de pate cuite sur une plaque et servi avec du sirop ?", 'da': 'Hvilken morgenmad er lavet af rørt dej tilberedt på en pande og serveret med sirup?'},
 'opt1': {'en': 'Pancakes', 'de': 'Pfannkuchen', 'es': 'Panqueques', 'nl': 'Pannenkoeken', 'sv': 'Pannkakor', 'fr': 'Pancakes', 'da': 'Pandekager'},
 'opt2': {'en': 'Waffles', 'de': 'Waffeln', 'es': 'Gofres', 'nl': 'Wafels', 'sv': 'Vafflor', 'fr': 'Gaufres', 'da': 'Vafler'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which ship famously sank after hitting an iceberg in 1912?', 'de': 'Welches Schiff sank beruehmt nach einer Kollision mit einem Eisberg 1912?', 'es': '¿Que barco se hundio famosamente tras chocar con un iceberg en 1912?', 'nl': 'Welk schip zonk beroemd na een aanvaring met een ijsberg in 1912?', 'sv': 'Vilket fartyg sjonk beromt efter att ha kolliderat med ett isberg 1912?', 'fr': "Quel navire a coule celebrement apres avoir heurte un iceberg en 1912 ?", 'da': 'Hvilket skib sank berømt efter at have ramt et isbjerg i 1912?'},
 'opt1': {'en': 'Titanic', 'de': 'Titanic', 'es': 'Titanic', 'nl': 'Titanic', 'sv': 'Titanic', 'fr': 'Titanic', 'da': 'Titanic'},
 'opt2': {'en': 'Lusitania', 'de': 'Lusitania', 'es': 'Lusitania', 'nl': 'Lusitania', 'sv': 'Lusitania', 'fr': 'Lusitania', 'da': 'Lusitania'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which decade is associated with the birth of rock and roll, with artists like Elvis Presley?', 'de': 'Mit welchem Jahrzehnt wird die Geburt von Rock and Roll mit Kuenstlern wie Elvis Presley assoziiert?', 'es': '¿Con que decada se asocia el nacimiento del rock and roll, con artistas como Elvis Presley?', 'nl': "Met welk decennium wordt de geboorte van rock-'n-roll geassocieerd, met artiesten als Elvis Presley?", 'sv': 'Vilket artionde forknippas med rock and rolls fodelse, med artister som Elvis Presley?', 'fr': "A quelle decennie associe-t-on la naissance du rock and roll, avec des artistes comme Elvis Presley ?", 'da': 'Hvilket årti forbindes med fødslen af rock and roll, med kunstnere som Elvis Presley?'},
 'opt1': {'en': '1950s', 'de': '1950er', 'es': 'los años 50', 'nl': 'jaren 50', 'sv': '1950-talet', 'fr': 'les annees 1950', 'da': '1950’erne'},
 'opt2': {'en': '1960s', 'de': '1960er', 'es': 'los años 60', 'nl': 'jaren 60', 'sv': '1960-talet', 'fr': 'les annees 1960', 'da': '1960’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which country is shaped like a long, thin strip along South America's western coast?", 'de': 'Welches Land ist wie ein langer, duenner Streifen entlang der Westkueste Suedamerikas geformt?', 'es': '¿Que pais tiene forma de franja larga y delgada a lo largo de la costa occidental de Sudamerica?', 'nl': 'Welk land heeft de vorm van een lange, dunne strook langs de westkust van Zuid-Amerika?', 'sv': 'Vilket land ar format som en lang, smal remsa langs Sydamerikas vastkust?', 'fr': "Quel pays a la forme d'une longue bande etroite le long de la cote ouest de l'Amerique du Sud ?", 'da': 'Hvilket land har form som en lang, tynd stribe langs Sydamerikas vestkyst?'},
 'opt1': {'en': 'Chile', 'de': 'Chile', 'es': 'Chile', 'nl': 'Chili', 'sv': 'Chile', 'fr': 'le Chili', 'da': 'Chile'},
 'opt2': {'en': 'Peru', 'de': 'Peru', 'es': 'Peru', 'nl': 'Peru', 'sv': 'Peru', 'fr': 'le Perou', 'da': 'Peru'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which marine animal has five arms and can regenerate them?', 'de': 'Welches Meerestier hat fuenf Arme und kann sie regenerieren?', 'es': '¿Que animal marino tiene cinco brazos y puede regenerarlos?', 'nl': 'Welk zeedier heeft vijf armen en kan deze regenereren?', 'sv': 'Vilket havsdjur har fem armar och kan aterskapa dem?', 'fr': "Quel animal marin a cinq bras et peut les regenerer ?", 'da': 'Hvilket havdyr har fem arme og kan gendanne dem?'},
 'opt1': {'en': 'Starfish', 'de': 'Seestern', 'es': 'Estrella de mar', 'nl': 'Zeester', 'sv': 'Sjostjarna', 'fr': 'Etoile de mer', 'da': 'Søstjerne'},
 'opt2': {'en': 'Sea urchin', 'de': 'Seeigel', 'es': 'Erizo de mar', 'nl': 'Zee-egel', 'sv': 'Sjoborre', 'fr': 'Oursin', 'da': 'Søpindsvin'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which gemstone is famously blue and associated with the Titanic's 'Heart of the Ocean'?", 'de': "Welcher Edelstein ist beruehmt blau und wird mit dem 'Herz des Ozeans' aus Titanic assoziiert?", 'es': "¿Que piedra preciosa es famosamente azul y se asocia con el 'Corazon del Oceano' de Titanic?", 'nl': "Welke edelsteen is beroemd blauw en wordt geassocieerd met 'Hart van de Oceaan' uit Titanic?", 'sv': "Vilken adelsten ar beromt bla och forknippas med 'Havets hjarta' fran Titanic?", 'fr': "Quelle pierre precieuse est celebrement bleue et associee au 'Coeur de l'Ocean' du Titanic ?", 'da': "Hvilken ædelsten er berømt blå og forbundet med 'Havets Hjerte' fra Titanic?"},
 'opt1': {'en': 'Sapphire', 'de': 'Saphir', 'es': 'Zafiro', 'nl': 'Saffier', 'sv': 'Safir', 'fr': 'Saphir', 'da': 'Safir'},
 'opt2': {'en': 'Aquamarine', 'de': 'Aquamarin', 'es': 'Aguamarina', 'nl': 'Aquamarijn', 'sv': 'Akvamarin', 'fr': 'Aigue-marine', 'da': 'Akvamarin'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic toy is a plastic brick used to build structures?', 'de': 'Welches klassische Spielzeug ist ein Kunststoffziegel zum Bauen von Strukturen?', 'es': '¿Que juguete clasico es un ladrillo de plastico usado para construir estructuras?', 'nl': 'Welk klassiek speelgoed is een plastic blokje om structuren mee te bouwen?', 'sv': 'Vilken klassisk leksak ar en plastbricka som anvands for att bygga strukturer?', 'fr': "Quel jouet classique est une brique en plastique utilisee pour construire des structures ?", 'da': 'Hvilket klassisk legetøj er en plastikklods brugt til at bygge strukturer?'},
 'opt1': {'en': 'LEGO', 'de': 'LEGO', 'es': 'LEGO', 'nl': 'LEGO', 'sv': 'LEGO', 'fr': 'LEGO', 'da': 'LEGO'},
 'opt2': {'en': "K'Nex", 'de': "K'Nex", 'es': "K'Nex", 'nl': "K'Nex", 'sv': "K'Nex", 'fr': "K'Nex", 'da': "K'Nex"}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which nut is used to make traditional marzipan?', 'de': 'Welche Nuss wird fuer traditionelles Marzipan verwendet?', 'es': '¿Que fruto seco se usa para hacer el mazapan tradicional?', 'nl': 'Welke noot wordt gebruikt om traditionele marsepein te maken?', 'sv': 'Vilken nöt anvands for att gora traditionell marsipan?', 'fr': "Quel fruit a coque est utilise pour faire le massepain traditionnel ?", 'da': 'Hvilken nød bruges til at lave traditionel marcipan?'},
 'opt1': {'en': 'Almond', 'de': 'Mandel', 'es': 'Almendra', 'nl': 'Amandel', 'sv': 'Mandel', 'fr': 'Amande', 'da': 'Mandel'},
 'opt2': {'en': 'Walnut', 'de': 'Walnuss', 'es': 'Nuez', 'nl': 'Walnoot', 'sv': 'Valnot', 'fr': 'Noix', 'da': 'Valnød'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which flower is traditionally given on the 1st wedding anniversary in some traditions?', 'de': 'Welche Blume wird traditionell zum 1. Hochzeitstag in manchen Traditionen verschenkt?', 'es': '¿Que flor se regala tradicionalmente en el primer aniversario de boda en algunas tradiciones?', 'nl': 'Welke bloem wordt traditioneel gegeven op de eerste huwelijksverjaardag in sommige tradities?', 'sv': 'Vilken blomma ges traditionellt pa den forsta bröllopsdagen i vissa traditioner?', 'fr': "Quelle fleur est traditionnellement offerte pour le 1er anniversaire de mariage selon certaines traditions ?", 'da': 'Hvilken blomst gives traditionelt på den 1. bryllupsdag i nogle traditioner?'},
 'opt1': {'en': 'Carnation', 'de': 'Nelke', 'es': 'Clavel', 'nl': 'Anjer', 'sv': 'Nejlika', 'fr': 'Oeillet', 'da': 'Nellike'},
 'opt2': {'en': 'Lily', 'de': 'Lilie', 'es': 'Lirio', 'nl': 'Lelie', 'sv': 'Lilja', 'fr': 'Lys', 'da': 'Lilje'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to Yellowstone National Park (mostly)?', 'de': 'In welchem US-Bundesstaat liegt (groesstenteils) der Yellowstone-Nationalpark?', 'es': '¿En que estado de EE.UU. se encuentra (mayormente) el Parque Nacional Yellowstone?', 'nl': 'In welke Amerikaanse staat ligt (grotendeels) Yellowstone National Park?', 'sv': 'I vilken amerikansk delstat ligger (mestadels) Yellowstone nationalpark?', 'fr': "Dans quel État americain se trouve (surtout) le parc national de Yellowstone ?", 'da': 'I hvilken amerikansk delstat ligger (mest) Yellowstone nationalpark?'},
 'opt1': {'en': 'Wyoming', 'de': 'Wyoming', 'es': 'Wyoming', 'nl': 'Wyoming', 'sv': 'Wyoming', 'fr': 'le Wyoming', 'da': 'Wyoming'},
 'opt2': {'en': 'Montana', 'de': 'Montana', 'es': 'Montana', 'nl': 'Montana', 'sv': 'Montana', 'fr': 'le Montana', 'da': 'Montana'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves weaving fabric on a loom?', 'de': 'Bei welchem Handwerk wird Stoff auf einem Webstuhl gewebt?', 'es': '¿Que manualidad implica tejer tela en un telar?', 'nl': 'Bij welk handwerk wordt stof geweven op een weefgetouw?', 'sv': 'Vilket hantverk innebar att vava tyg pa en vavstol?', 'fr': "Quel art manuel consiste a tisser du tissu sur un metier a tisser ?", 'da': 'Hvilket håndværk involverer vævning af stof på en væv?'},
 'opt1': {'en': 'Weaving', 'de': 'Weben', 'es': 'Tejido en telar', 'nl': 'Weven', 'sv': 'Vavning', 'fr': 'Le tissage', 'da': 'Vævning'},
 'opt2': {'en': 'Knitting', 'de': 'Stricken', 'es': 'Tejido de punto', 'nl': 'Breien', 'sv': 'Stickning', 'fr': 'Le tricot', 'da': 'Strikning'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves rolling a heavy ball to knock down ten pins?', 'de': 'Bei welcher Sportart rollt man eine schwere Kugel, um zehn Pins umzuwerfen?', 'es': '¿Que deporte consiste en hacer rodar una bola pesada para derribar diez bolos?', 'nl': 'Bij welke sport rol je een zware bal om tien kegels omver te gooien?', 'sv': 'Vilken sport innebar att man rullar ett tungt klot for att fella tio kaglor?', 'fr': "Quel sport consiste a faire rouler une lourde boule pour renverser dix quilles ?", 'da': 'Hvilken sport indebærer at rulle en tung kugle for at vælte ti kegler?'},
 'opt1': {'en': 'Bowling', 'de': 'Bowling', 'es': 'Bolos', 'nl': 'Bowlen', 'sv': 'Bowling', 'fr': 'Bowling', 'da': 'Bowling'},
 'opt2': {'en': 'Bocce', 'de': 'Boccia', 'es': 'Petanca', 'nl': 'Jeu de boules', 'sv': 'Boule', 'fr': 'La petanque', 'da': 'Boccia'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which item of clothing is worn on the feet, usually with laces?', 'de': 'Welches Kleidungsstueck wird an den Fuessen getragen, meist mit Schnuersenkeln?', 'es': '¿Que prenda de vestir se lleva en los pies, generalmente con cordones?', 'nl': 'Welk kledingstuk wordt aan de voeten gedragen, meestal met veters?', 'sv': 'Vilket klädesplagg baers pa fotterna, vanligtvis med snorningar?', 'fr': "Quel vetement se porte aux pieds, generalement avec des lacets ?", 'da': 'Hvilket beklædningsgenstand bæres på fødderne, normalt med snørebånd?'},
 'opt1': {'en': 'Shoes', 'de': 'Schuhe', 'es': 'Zapatos', 'nl': 'Schoenen', 'sv': 'Skor', 'fr': 'Chaussures', 'da': 'Sko'},
 'opt2': {'en': 'Socks', 'de': 'Socken', 'es': 'Calcetines', 'nl': 'Sokken', 'sv': 'Strumpor', 'fr': 'Chaussettes', 'da': 'Sokker'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic show features puppets including Kermit the Frog?', 'de': 'In welcher klassischen Show treten Puppen wie Kermit der Frosch auf?', 'es': '¿Que programa clasico presenta a marionetas incluyendo a la Rana Rene?', 'nl': 'In welke klassieke show treden poppen op waaronder Kermit de Kikker?', 'sv': 'I vilken klassisk show forekommer dockor inklusive Kermit grodan?', 'fr': "Quelle emission classique met en scene des marionnettes dont Kermit la Grenouille ?", 'da': 'Hvilket klassisk show har dukker inklusive Kermit Frøen?'},
 'opt1': {'en': 'The Muppet Show', 'de': 'Die Muppet Show', 'es': 'Los Muppets', 'nl': 'The Muppet Show', 'sv': 'Mupparna', 'fr': 'Le Muppet Show', 'da': 'Muppet Show'},
 'opt2': {'en': 'Sesame Street', 'de': 'Sesamstraße', 'es': 'Barrio Sesamo', 'nl': 'Sesamstraat', 'sv': 'Sesam', 'fr': 'Rue Sesame', 'da': 'Sesamgade'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which European country is home to the city of Amsterdam?', 'de': 'In welchem europaeischen Land liegt die Stadt Amsterdam?', 'es': '¿En que pais europeo se encuentra la ciudad de Amsterdam?', 'nl': 'In welk Europees land ligt de stad Amsterdam?', 'sv': 'I vilket europeiskt land ligger staden Amsterdam?', 'fr': "Dans quel pays europeen se trouve la ville d'Amsterdam ?", 'da': 'Hvilket europæisk land huser byen Amsterdam?'},
 'opt1': {'en': 'Netherlands', 'de': 'Niederlande', 'es': 'Paises Bajos', 'nl': 'Nederland', 'sv': 'Nederlanderna', 'fr': 'les Pays-Bas', 'da': 'Nederlandene'},
 'opt2': {'en': 'Belgium', 'de': 'Belgien', 'es': 'Belgica', 'nl': 'Belgie', 'sv': 'Belgien', 'fr': 'la Belgique', 'da': 'Belgien'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was the first to have a professional career (as a lawyer) before the White House?', 'de': 'Welche First Lady hatte als erste eine berufliche Karriere (als Anwaltin) vor dem Weißen Haus?', 'es': '¿Que primera dama fue la primera en tener una carrera profesional (como abogada) antes de la Casa Blanca?', 'nl': 'Welke First Lady had als eerste een professionele carriere (als advocaat) voor het Witte Huis?', 'sv': 'Vilken First Lady var forst med en yrkeskarriar (som advokat) fore Vita huset?', 'fr': "Quelle Première dame a ete la premiere a avoir une carriere professionnelle (avocate) avant la Maison-Blanche ?", 'da': 'Hvilken førstedame var den første til at have en professionel karriere (som advokat) før Det Hvide Hus?'},
 'opt1': {'en': 'Hillary Clinton', 'de': 'Hillary Clinton', 'es': 'Hillary Clinton', 'nl': 'Hillary Clinton', 'sv': 'Hillary Clinton', 'fr': 'Hillary Clinton', 'da': 'Hillary Clinton'},
 'opt2': {'en': 'Michelle Obama', 'de': 'Michelle Obama', 'es': 'Michelle Obama', 'nl': 'Michelle Obama', 'sv': 'Michelle Obama', 'fr': 'Michelle Obama', 'da': 'Michelle Obama'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Mexican dish consists of corn dough filled with meat and steamed in a husk?', 'de': 'Welches mexikanische Gericht besteht aus Maisteig, gefuellt mit Fleisch und in einer Huelle gedaempft?', 'es': '¿Que plato mexicano consiste en masa de maiz rellena de carne y cocida al vapor en una hoja?', 'nl': 'Welk Mexicaans gerecht bestaat uit maisdeeg gevuld met vlees en gestoomd in een schil?', 'sv': 'Vilken mexikansk ratt bestar av majsdeg fylld med kott och angkokt i ett skal?', 'fr': "Quel plat mexicain consiste en une pate de mais fourree de viande et cuite a la vapeur dans une enveloppe ?", 'da': 'Hvilken mexicansk ret består af majsdej fyldt med kød og dampet i et skal?'},
 'opt1': {'en': 'Tamale', 'de': 'Tamale', 'es': 'Tamal', 'nl': 'Tamale', 'sv': 'Tamale', 'fr': 'Tamale', 'da': 'Tamale'},
 'opt2': {'en': 'Enchilada', 'de': 'Enchilada', 'es': 'Enchilada', 'nl': 'Enchilada', 'sv': 'Enchilada', 'fr': 'Enchilada', 'da': 'Enchilada'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which part of the brain controls balance and coordination?', 'de': 'Welcher Teil des Gehirns steuert Gleichgewicht und Koordination?', 'es': '¿Que parte del cerebro controla el equilibrio y la coordinacion?', 'nl': 'Welk deel van de hersenen regelt balans en coordinatie?', 'sv': 'Vilken del av hjarnan styr balans och koordination?', 'fr': "Quelle partie du cerveau controle l'equilibre et la coordination ?", 'da': 'Hvilken del af hjernen styrer balance og koordination?'},
 'opt1': {'en': 'Cerebellum', 'de': 'Kleinhirn', 'es': 'Cerebelo', 'nl': 'Kleine hersenen', 'sv': 'Lillhjarnan', 'fr': 'Cervelet', 'da': 'Lillehjerne'},
 'opt2': {'en': 'Cerebrum', 'de': 'Großhirn', 'es': 'Cerebro', 'nl': 'Grote hersenen', 'sv': 'Storhjarnan', 'fr': 'Cerveau', 'da': 'Storhjerne'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which author wrote 'To the Lighthouse'?", 'de': "Welche Autorin schrieb 'Die Fahrt zum Leuchtturm'?", 'es': "¿Que autora escribio 'Al faro'?", 'nl': "Welke auteur schreef 'Naar de vuurtoren'?", 'sv': "Vilken forfattare skrev 'Mot fyren'?", 'fr': "Quelle autrice a ecrit 'La Promenade au phare' ?", 'da': "Hvilken forfatter skrev 'Til fyret'?"},
 'opt1': {'en': 'Virginia Woolf', 'de': 'Virginia Woolf', 'es': 'Virginia Woolf', 'nl': 'Virginia Woolf', 'sv': 'Virginia Woolf', 'fr': 'Virginia Woolf', 'da': 'Virginia Woolf'},
 'opt2': {'en': 'Edith Wharton', 'de': 'Edith Wharton', 'es': 'Edith Wharton', 'nl': 'Edith Wharton', 'sv': 'Edith Wharton', 'fr': 'Edith Wharton', 'da': 'Edith Wharton'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which actress played Bridget Jones in the film series?', 'de': 'Welche Schauspielerin spielte Bridget Jones in der Filmreihe?', 'es': '¿Que actriz interpreto a Bridget Jones en la serie de peliculas?', 'nl': 'Welke actrice speelde Bridget Jones in de filmreeks?', 'sv': 'Vilken skadespelerska spelade Bridget Jones i filmserien?', 'fr': "Quelle actrice a joue Bridget Jones dans la serie de films ?", 'da': 'Hvilken skuespiller spillede Bridget Jones i filmserien?'},
 'opt1': {'en': 'Renee Zellweger', 'de': 'Renee Zellweger', 'es': 'Renee Zellweger', 'nl': 'Renee Zellweger', 'sv': 'Renee Zellweger', 'fr': 'Renee Zellweger', 'da': 'Renee Zellweger'},
 'opt2': {'en': 'Kate Winslet', 'de': 'Kate Winslet', 'es': 'Kate Winslet', 'nl': 'Kate Winslet', 'sv': 'Kate Winslet', 'fr': 'Kate Winslet', 'da': 'Kate Winslet'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which singer is known as 'The Boss'?", 'de': "Welcher Saenger wird 'The Boss' genannt?", 'es': "¿Que cantante es conocido como 'The Boss'?", 'nl': "Welke zanger staat bekend als 'The Boss'?", 'sv': "Vilken sangare kallas 'The Boss'?", 'fr': "Quel chanteur est surnomme 'The Boss' ?", 'da': "Hvilken sanger er kendt som 'The Boss'?"},
 'opt1': {'en': 'Bruce Springsteen', 'de': 'Bruce Springsteen', 'es': 'Bruce Springsteen', 'nl': 'Bruce Springsteen', 'sv': 'Bruce Springsteen', 'fr': 'Bruce Springsteen', 'da': 'Bruce Springsteen'},
 'opt2': {'en': 'Billy Joel', 'de': 'Billy Joel', 'es': 'Billy Joel', 'nl': 'Billy Joel', 'sv': 'Billy Joel', 'fr': 'Billy Joel', 'da': 'Billy Joel'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to Machu Picchu, the ancient Inca citadel?', 'de': 'In welchem Land liegt Machu Picchu, die antike Inka-Zitadelle?', 'es': '¿En que pais esta Machu Picchu, la antigua ciudadela inca?', 'nl': 'In welk land ligt Machu Picchu, de oude Inca-citadel?', 'sv': 'I vilket land ligger Machu Picchu, den antika inkacitadellen?', 'fr': "Dans quel pays se trouve le Machu Picchu, l'ancienne citadelle inca ?", 'da': 'Hvilket land huser Machu Picchu, den gamle inkaborg?'},
 'opt1': {'en': 'Peru', 'de': 'Peru', 'es': 'Peru', 'nl': 'Peru', 'sv': 'Peru', 'fr': 'le Perou', 'da': 'Peru'},
 'opt2': {'en': 'Bolivia', 'de': 'Bolivien', 'es': 'Bolivia', 'nl': 'Bolivia', 'sv': 'Bolivia', 'fr': 'la Bolivie', 'da': 'Bolivia'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large African mammal has the longest neck of any animal?', 'de': 'Welches große afrikanische Saeugetier hat den laengsten Hals aller Tiere?', 'es': '¿Que gran mamifero africano tiene el cuello mas largo de cualquier animal?', 'nl': 'Welk groot Afrikaans zoogdier heeft de langste nek van alle dieren?', 'sv': 'Vilket stort afrikanskt daggdjur har den langsta halsen av alla djur?', 'fr': "Quel grand mammifere africain a le cou le plus long de tous les animaux ?", 'da': 'Hvilket stort afrikansk pattedyr har den længste hals blandt alle dyr?'},
 'opt1': {'en': 'Giraffe', 'de': 'Giraffe', 'es': 'Jirafa', 'nl': 'Giraf', 'sv': 'Giraff', 'fr': 'Girafe', 'da': 'Giraf'},
 'opt2': {'en': 'Camel', 'de': 'Kamel', 'es': 'Camello', 'nl': 'Kameel', 'sv': 'Kamel', 'fr': 'Chameau', 'da': 'Kamel'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Indian bread is cooked in a clay tandoor oven?', 'de': 'Welches indische Brot wird in einem Lehmofen (Tandoor) gebacken?', 'es': '¿Que pan indio se cocina en un horno de barro tandoor?', 'nl': 'Welk Indiaas brood wordt gebakken in een kleien tandoor-oven?', 'sv': 'Vilket indiskt brod tillagas i en lerugn (tandoor)?', 'fr': "Quel pain indien est cuit dans un four en argile appele tandoor ?", 'da': 'Hvilket indisk brød tilberedes i en lerovn (tandoor)?'},
 'opt1': {'en': 'Naan', 'de': 'Naan', 'es': 'Naan', 'nl': 'Naan', 'sv': 'Naan', 'fr': 'Naan', 'da': 'Naan'},
 'opt2': {'en': 'Roti', 'de': 'Roti', 'es': 'Roti', 'nl': 'Roti', 'sv': 'Roti', 'fr': 'Roti', 'da': 'Roti'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which First Lady was known for her 'Beautification of America' highway campaign?", 'de': "Welche First Lady war bekannt fuer ihre 'Verschoenerung Amerikas'-Autobahnkampagne?", 'es': "¿Que primera dama fue conocida por su campaña de 'Embellecimiento de America' en las carreteras?", 'nl': "Welke First Lady stond bekend om haar 'Beautification of America' snelwegcampagne?", 'sv': "Vilken First Lady var kand for sin 'Beautification of America'-vagkampanj?", 'fr': "Quelle Première dame etait connue pour sa campagne autoroutiere 'Beautification of America' ?", 'da': "Hvilken førstedame var kendt for sin 'Beautification of America' motorvejskampagne?"},
 'opt1': {'en': 'Lady Bird Johnson', 'de': 'Lady Bird Johnson', 'es': 'Lady Bird Johnson', 'nl': 'Lady Bird Johnson', 'sv': 'Lady Bird Johnson', 'fr': 'Lady Bird Johnson', 'da': 'Lady Bird Johnson'},
 'opt2': {'en': 'Pat Nixon', 'de': 'Pat Nixon', 'es': 'Pat Nixon', 'nl': 'Pat Nixon', 'sv': 'Pat Nixon', 'fr': 'Pat Nixon', 'da': 'Pat Nixon'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which artist is known for melting clock paintings like 'The Persistence of Memory'?", 'de': "Welcher Kuenstler ist fuer schmelzende Uhr-Gemaelde wie 'Die Bestaendigkeit der Erinnerung' bekannt?", 'es': "¿Que artista es conocido por pinturas de relojes derretidos como 'La persistencia de la memoria'?", 'nl': "Welke kunstenaar staat bekend om smeltende klokschilderijen zoals 'De volharding der herinnering'?", 'sv': "Vilken konstnar ar kand for smaltande klockmalningar som 'Minnets envishet'?", 'fr': "Quel artiste est connu pour ses peintures d'horloges fondantes comme 'La Persistance de la memoire' ?", 'da': "Hvilken kunstner er kendt for smeltende ur-malerier som 'Erindringens vedholdenhed'?"},
 'opt1': {'en': 'Salvador Dali', 'de': 'Salvador Dali', 'es': 'Salvador Dali', 'nl': 'Salvador Dali', 'sv': 'Salvador Dali', 'fr': 'Salvador Dali', 'da': 'Salvador Dali'},
 'opt2': {'en': 'Rene Magritte', 'de': 'Rene Magritte', 'es': 'Rene Magritte', 'nl': 'Rene Magritte', 'sv': 'Rene Magritte', 'fr': 'Rene Magritte', 'da': 'Rene Magritte'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which tree produces edible seeds inside cones and is a symbol of Christmas?', 'de': 'Welcher Baum produziert essbare Samen in Zapfen und ist ein Symbol fuer Weihnachten?', 'es': '¿Que arbol produce semillas comestibles dentro de conos y es un simbolo de la Navidad?', 'nl': 'Welke boom produceert eetbare zaden in kegels en is een symbool van Kerstmis?', 'sv': 'Vilket trad producerar atbara fron i kottar och ar en julsymbol?', 'fr': "Quel arbre produit des graines comestibles dans des cones et est un symbole de Noel ?", 'da': 'Hvilket træ producerer spiselige frø i kogler og er et julesymbol?'},
 'opt1': {'en': 'Fir', 'de': 'Tanne', 'es': 'Abeto', 'nl': 'Den', 'sv': 'Gran', 'fr': 'Sapin', 'da': 'Gran'},
 'opt2': {'en': 'Spruce', 'de': 'Fichte', 'es': 'Picea', 'nl': 'Spar', 'sv': 'Gran (spruce)', 'fr': 'Epicea', 'da': 'Rødgran'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which musical is based on the life of Alexander Hamilton?', 'de': 'Welches Musical basiert auf dem Leben von Alexander Hamilton?', 'es': '¿Que musical se basa en la vida de Alexander Hamilton?', 'nl': 'Welke musical is gebaseerd op het leven van Alexander Hamilton?', 'sv': 'Vilken musikal ar baserad pa Alexander Hamiltons liv?', 'fr': "Quelle comedie musicale s'inspire de la vie d'Alexander Hamilton ?", 'da': 'Hvilken musical er baseret på Alexander Hamiltons liv?'},
 'opt1': {'en': 'Hamilton', 'de': 'Hamilton', 'es': 'Hamilton', 'nl': 'Hamilton', 'sv': 'Hamilton', 'fr': 'Hamilton', 'da': 'Hamilton'},
 'opt2': {'en': '1776', 'de': '1776', 'es': '1776', 'nl': '1776', 'sv': '1776', 'fr': '1776', 'da': '1776'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country has Vienna as its capital?', 'de': 'Welches Land hat Wien als Hauptstadt?', 'es': '¿Que pais tiene a Viena como su capital?', 'nl': 'Welk land heeft Wenen als hoofdstad?', 'sv': 'Vilket land har Wien som huvudstad?', 'fr': "Quel pays a Vienne pour capitale ?", 'da': 'Hvilket land har Wien som hovedstad?'},
 'opt1': {'en': 'Austria', 'de': 'Österreich', 'es': 'Austria', 'nl': 'Oostenrijk', 'sv': 'Osterrike', 'fr': "l'Autriche", 'da': 'Østrig'},
 'opt2': {'en': 'Switzerland', 'de': 'Schweiz', 'es': 'Suiza', 'nl': 'Zwitserland', 'sv': 'Schweiz', 'fr': 'la Suisse', 'da': 'Schweiz'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which spice is ground from dried chili peppers and used to season Cajun food?', 'de': 'Welches Gewürz wird aus getrockneten Chilischoten gemahlen und fuer Cajun-Speisen verwendet?', 'es': '¿Que especia se muele de chiles secos y se usa para sazonar la comida cajun?', 'nl': 'Welke specerij wordt gemalen van gedroogde chilipepers en gebruikt om Cajun-gerechten te kruiden?', 'sv': 'Vilken krydda mals fran torkade chilipeppar och anvands for att smaksatta cajunmat?', 'fr': "Quelle epice est moulue a partir de piments seches et utilisee pour assaisonner la cuisine cajun ?", 'da': 'Hvilket krydderi males af tørrede chilipeber og bruges til at krydre cajun-mad?'},
 'opt1': {'en': 'Cayenne pepper', 'de': 'Cayennepfeffer', 'es': 'Pimienta de cayena', 'nl': 'Cayennepeper', 'sv': 'Cayennepeppar', 'fr': 'Poivre de Cayenne', 'da': 'Cayennepeber'},
 'opt2': {'en': 'Black pepper', 'de': 'Schwarzer Pfeffer', 'es': 'Pimienta negra', 'nl': 'Zwarte peper', 'sv': 'Svartpeppar', 'fr': 'Poivre noir', 'da': 'Sort peber'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which needlework technique uses a needle to create pictures with small X-shaped stitches?', 'de': 'Welche Handarbeitstechnik verwendet eine Nadel, um Bilder mit kleinen X-foermigen Stichen zu erstellen?', 'es': '¿Que tecnica de costura usa una aguja para crear imagenes con pequeñas puntadas en forma de X?', 'nl': 'Welke naaitechniek gebruikt een naald om afbeeldingen te maken met kleine X-vormige steken?', 'sv': 'Vilken handarbetsteknik anvander en nal for att skapa bilder med sma X-formade stygn?', 'fr': "Quelle technique de couture utilise une aiguille pour creer des images avec de petits points en X ?", 'da': 'Hvilken håndarbejdsteknik bruger en nål til at skabe billeder med små X-formede sting?'},
 'opt1': {'en': 'Cross-stitch', 'de': 'Kreuzstich', 'es': 'Punto de cruz', 'nl': 'Kruissteek', 'sv': 'Korsstygn', 'fr': 'Point de croix', 'da': 'Korssting'},
 'opt2': {'en': 'Embroidery', 'de': 'Stickerei', 'es': 'Bordado', 'nl': 'Borduurwerk', 'sv': 'Broderi', 'fr': 'Broderie', 'da': 'Broderi'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which bone is the longest in the human body?', 'de': 'Welcher Knochen ist der laengste im menschlichen Koerper?', 'es': '¿Que hueso es el mas largo del cuerpo humano?', 'nl': 'Welk bot is het langste in het menselijk lichaam?', 'sv': 'Vilket ben ar langst i den manskliga kroppen?', 'fr': "Quel os est le plus long du corps humain ?", 'da': 'Hvilket knogle er den længste i den menneskelige krop?'},
 'opt1': {'en': 'Femur', 'de': 'Oberschenkelknochen', 'es': 'Femur', 'nl': 'Dijbeen', 'sv': 'Lårben', 'fr': 'Femur', 'da': 'Lårben'},
 'opt2': {'en': 'Tibia', 'de': 'Schienbein', 'es': 'Tibia', 'nl': 'Scheenbeen', 'sv': 'Skenben', 'fr': 'Tibia', 'da': 'Skinneben'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'In figure skating, what is a jump with four rotations called?', 'de': 'Wie nennt man beim Eiskunstlauf einen Sprung mit vier Umdrehungen?', 'es': '¿Como se llama en patinaje artistico un salto con cuatro rotaciones?', 'nl': 'Hoe heet bij kunstschaatsen een sprong met vier rotaties?', 'sv': 'Vad kallas ett hopp med fyra rotationer i konståkning?', 'fr': "En patinage artistique, comment appelle-t-on un saut avec quatre rotations ?", 'da': 'Hvad kaldes et spring med fire rotationer i kunstskøjteløb?'},
 'opt1': {'en': 'Quad', 'de': 'Vierfachsprung', 'es': 'Cuadruple', 'nl': 'Quad', 'sv': 'Quad', 'fr': 'Quadruple saut', 'da': 'Quad'},
 'opt2': {'en': 'Axel', 'de': 'Axel', 'es': 'Axel', 'nl': 'Axel', 'sv': 'Axel', 'fr': 'Axel', 'da': 'Axel'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Italian city is home to the Uffizi Gallery and the Duomo cathedral?', 'de': 'In welcher italienischen Stadt befinden sich die Uffizien und der Dom?', 'es': '¿En que ciudad italiana estan la Galeria Uffizi y la catedral del Duomo?', 'nl': 'In welke Italiaanse stad bevinden zich de Uffizi-galerij en de Duomo?', 'sv': 'I vilken italiensk stad ligger Uffizierna och Duomon?', 'fr': "Dans quelle ville italienne se trouvent la galerie des Offices et la cathedrale du Duomo ?", 'da': 'I hvilken italiensk by ligger Uffizigalleriet og Duomo-katedralen?'},
 'opt1': {'en': 'Florence', 'de': 'Florenz', 'es': 'Florencia', 'nl': 'Florence', 'sv': 'Florens', 'fr': 'Florence', 'da': 'Firenze'},
 'opt2': {'en': 'Venice', 'de': 'Venedig', 'es': 'Venecia', 'nl': 'Venetie', 'sv': 'Venedig', 'fr': 'Venise', 'da': 'Venedig'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Rose in the film 'Titanic'?", 'de': "Welche Schauspielerin spielte Rose im Film 'Titanic'?", 'es': "¿Que actriz interpreto a Rose en la pelicula 'Titanic'?", 'nl': "Welke actrice speelde Rose in de film 'Titanic'?", 'sv': "Vilken skadespelerska spelade Rose i filmen 'Titanic'?", 'fr': "Quelle actrice a joue Rose dans le film 'Titanic' ?", 'da': "Hvilken skuespiller spillede Rose i filmen 'Titanic'?"},
 'opt1': {'en': 'Kate Winslet', 'de': 'Kate Winslet', 'es': 'Kate Winslet', 'nl': 'Kate Winslet', 'sv': 'Kate Winslet', 'fr': 'Kate Winslet', 'da': 'Kate Winslet'},
 'opt2': {'en': 'Gwyneth Paltrow', 'de': 'Gwyneth Paltrow', 'es': 'Gwyneth Paltrow', 'nl': 'Gwyneth Paltrow', 'sv': 'Gwyneth Paltrow', 'fr': 'Gwyneth Paltrow', 'da': 'Gwyneth Paltrow'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which classic novel follows Captain Nemo and the submarine Nautilus?', 'de': 'Welcher Klassiker handelt von Kapitaen Nemo und dem U-Boot Nautilus?', 'es': '¿Que novela clasica sigue al Capitan Nemo y el submarino Nautilus?', 'nl': 'Welke klassieke roman volgt Kapitein Nemo en de onderzeeer Nautilus?', 'sv': 'Vilken klassisk roman foljer Kapten Nemo och ubaten Nautilus?', 'fr': "Quel roman classique suit le capitaine Nemo et le sous-marin Nautilus ?", 'da': 'Hvilken klassisk roman følger Kaptajn Nemo og ubåden Nautilus?'},
 'opt1': {'en': 'Twenty Thousand Leagues Under the Sea', 'de': '20.000 Meilen unter dem Meer', 'es': 'Veinte mil leguas de viaje submarino', 'nl': 'Twintigduizend Mijlen onder Zee', 'sv': 'En vardsomsegling under havet', 'fr': 'Vingt Mille Lieues sous les mers', 'da': 'En Verdensomsejling under Havet'},
 'opt2': {'en': 'The Old Man and the Sea', 'de': 'Der alte Mann und das Meer', 'es': 'El viejo y el mar', 'nl': 'De oude man en de zee', 'sv': 'Den gamle och havet', 'fr': 'Le Vieil Homme et la Mer', 'da': 'Den gamle mand og havet'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which 1970s band is known for the album 'Rumours'?", 'de': "Welche Band der 1970er ist fuer das Album 'Rumours' bekannt?", 'es': "¿Que banda de los 70 es conocida por el album 'Rumours'?", 'nl': "Welke band uit de jaren 70 staat bekend om het album 'Rumours'?", 'sv': "Vilket 1970-talsband ar kant for albumet 'Rumours'?", 'fr': "Quel groupe des annees 1970 est connu pour l'album 'Rumours' ?", 'da': "Hvilket 1970’er-band er kendt for albummet 'Rumours'?"},
 'opt1': {'en': 'Fleetwood Mac', 'de': 'Fleetwood Mac', 'es': 'Fleetwood Mac', 'nl': 'Fleetwood Mac', 'sv': 'Fleetwood Mac', 'fr': 'Fleetwood Mac', 'da': 'Fleetwood Mac'},
 'opt2': {'en': 'The Eagles', 'de': 'The Eagles', 'es': 'The Eagles', 'nl': 'The Eagles', 'sv': 'The Eagles', 'fr': 'The Eagles', 'da': 'The Eagles'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is fried chicken served with waffles?', 'de': 'Welches suddstaatliche Gericht ist gebratenes Huhn mit Waffeln?', 'es': '¿Que plato sureño de EE.UU. es pollo frito servido con gofres?', 'nl': 'Welk zuidelijk Amerikaans gerecht is gefrituurde kip geserveerd met wafels?', 'sv': 'Vilken sydstatlig ratt ar friterad kyckling serverad med vafflor?', 'fr': "Quel plat du Sud des États-Unis est du poulet frit servi avec des gaufres ?", 'da': 'Hvilken sydstatsret er stegt kylling serveret med vafler?'},
 'opt1': {'en': 'Chicken and waffles', 'de': 'Chicken and Waffles', 'es': 'Pollo con gofres', 'nl': 'Kip met wafels', 'sv': 'Kyckling och vafflor', 'fr': 'Poulet et gaufres', 'da': 'Kylling og vafler'},
 'opt2': {'en': 'Fried chicken sandwich', 'de': 'Chicken-Sandwich', 'es': 'Sandwich de pollo frito', 'nl': 'Gefrituurde kipsandwich', 'sv': 'Friterad kycklingsmorgas', 'fr': 'Sandwich au poulet frit', 'da': 'Stegt kyllingesandwich'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which queen was known as the 'Virgin Queen' of England?", 'de': "Welche Koenigin wurde als 'Jungfraeuliche Koenigin' Englands bezeichnet?", 'es': "¿Que reina fue conocida como la 'Reina Virgen' de Inglaterra?", 'nl': "Welke koningin stond bekend als de 'Maagdelijke Koningin' van Engeland?", 'sv': "Vilken drottning kallades 'Jungfrudrottningen' av England?", 'fr': "Quelle reine etait connue comme la 'Reine Vierge' d'Angleterre ?", 'da': "Hvilken dronning var kendt som 'Jomfrudronningen' af England?"},
 'opt1': {'en': 'Elizabeth I', 'de': 'Elisabeth I.', 'es': 'Isabel I', 'nl': 'Elizabeth I', 'sv': 'Elizabeth I', 'fr': 'Elisabeth Ire', 'da': 'Elizabeth 1.'},
 'opt2': {'en': 'Mary I', 'de': 'Maria I.', 'es': 'Maria I', 'nl': 'Mary I', 'sv': 'Maria I', 'fr': 'Marie Ire', 'da': 'Maria 1.'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which North American amphibian can survive being partially frozen solid over winter?', 'de': 'Welches nordamerikanische Amphibium kann uberleben, wenn es im Winter teilweise gefriert?', 'es': '¿Que anfibio norteamericano puede sobrevivir estando parcialmente congelado durante el invierno?', 'nl': 'Welk Noord-Amerikaans amfibie kan overleven als het \'s winters gedeeltelijk bevriest?', 'sv': 'Vilket nordamerikanskt groddjur kan overleva att vara delvis fruset under vintern?', 'fr': "Quel amphibien nord-americain peut survivre en etant partiellement gele en hiver ?", 'da': 'Hvilket nordamerikansk padde kan overleve delvist frosset om vinteren?'},
 'opt1': {'en': 'Wood frog', 'de': 'Waldfrosch', 'es': 'Rana de bosque', 'nl': 'Bosskikker', 'sv': 'Skogsgroda', 'fr': 'Grenouille des bois', 'da': 'Skovfrø'},
 'opt2': {'en': 'Painted turtle', 'de': 'Zierschildkroete', 'es': 'Tortuga pintada', 'nl': 'Sierschildpad', 'sv': 'Malad sköldpadda', 'fr': 'Tortue peinte', 'da': 'Malet skildpadde'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the medieval walled city of Carcassonne?', 'de': 'In welchem Land liegt die mittelalterliche ummauerte Stadt Carcassonne?', 'es': '¿En que pais esta la ciudad amurallada medieval de Carcasona?', 'nl': 'In welk land ligt de middeleeuwse ommuurde stad Carcassonne?', 'sv': 'I vilket land ligger den medeltida murade staden Carcassonne?', 'fr': "Dans quel pays se trouve la cite medievale fortifiee de Carcassonne ?", 'da': 'Hvilket land huser den middelalderlige befæstede by Carcassonne?'},
 'opt1': {'en': 'France', 'de': 'Frankreich', 'es': 'Francia', 'nl': 'Frankrijk', 'sv': 'Frankrike', 'fr': 'la France', 'da': 'Frankrig'},
 'opt2': {'en': 'Spain', 'de': 'Spanien', 'es': 'España', 'nl': 'Spanje', 'sv': 'Spanien', 'fr': "l'Espagne", 'da': 'Spanien'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which type of handbag is named after Britain's Queen and carried on the arm?", 'de': "Welche Handtaschenart ist nach Großbritanniens Koenigin benannt und wird am Arm getragen?", 'es': "¿Que tipo de bolso lleva el nombre de la reina britanica y se lleva en el brazo?", 'nl': "Welk type handtas is genoemd naar de Britse koningin en wordt aan de arm gedragen?", 'sv': "Vilken typ av handvaska ar uppkallad efter Storbritanniens drottning och baers pa armen?", 'fr': "Quel type de sac a main porte le nom de la reine britannique et se porte au bras ?", 'da': 'Hvilken type håndtaske er opkaldt efter Storbritanniens dronning og bæres på armen?'},
 'opt1': {'en': 'Clutch', 'de': 'Clutch', 'es': 'Clutch', 'nl': 'Clutch', 'sv': 'Kuvertvaska', 'fr': 'Pochette', 'da': 'Clutch'},
 'opt2': {'en': 'Tote', 'de': 'Tote Bag', 'es': 'Bolso tote', 'nl': 'Tote bag', 'sv': 'Tygkasse', 'fr': 'Sac fourre-tout', 'da': 'Tote-taske'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Netflix series follows a chess prodigy named Beth Harmon?', 'de': 'Welche Netflix-Serie handelt von einem Schachwunderkind namens Beth Harmon?', 'es': '¿Que serie de Netflix sigue a un prodigio del ajedrez llamado Beth Harmon?', 'nl': 'Welke Netflix-serie volgt een schaakwonderkind genaamd Beth Harmon?', 'sv': 'Vilken Netflix-serie foljer ett schackgeni vid namn Beth Harmon?', 'fr': "Quelle serie Netflix suit un prodige des echecs nomme Beth Harmon ?", 'da': 'Hvilken Netflix-serie følger et skakgeni ved navn Beth Harmon?'},
 'opt1': {'en': "The Queen's Gambit", 'de': 'Das Damengambit', 'es': 'Gambito de dama', 'nl': "The Queen's Gambit", 'sv': 'Damgambit', 'fr': 'Le Jeu de la Dame', 'da': 'Dronninggambit'},
 'opt2': {'en': 'Bridgerton', 'de': 'Bridgerton', 'es': 'Bridgerton', 'nl': 'Bridgerton', 'sv': 'Bridgerton', 'fr': 'Bridgerton', 'da': 'Bridgerton'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which flower is the national flower of the Netherlands?', 'de': 'Welche Blume ist die Nationalblume der Niederlande?', 'es': '¿Que flor es la flor nacional de los Paises Bajos?', 'nl': 'Welke bloem is de nationale bloem van Nederland?', 'sv': 'Vilken blomma ar Nederlandernas nationalblomma?', 'fr': "Quelle fleur est la fleur nationale des Pays-Bas ?", 'da': 'Hvilken blomst er Nederlandenes nationalblomst?'},
 'opt1': {'en': 'Tulip', 'de': 'Tulpe', 'es': 'Tulipan', 'nl': 'Tulp', 'sv': 'Tulpan', 'fr': 'Tulipe', 'da': 'Tulipan'},
 'opt2': {'en': 'Daffodil', 'de': 'Osterglocke', 'es': 'Narciso', 'nl': 'Narcis', 'sv': 'Pasklilja', 'fr': 'Jonquille', 'da': 'Pinselilje'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Greek pastry is layered with phyllo dough, nuts, and honey syrup?', 'de': 'Welches griechische Gebaeck wird aus Schichten von Phyllo-Teig, Nuessen und Honigsirup hergestellt?', 'es': '¿Que pasteleria griega se hace con capas de masa filo, nueces y almibar de miel?', 'nl': 'Welk Grieks gebak bestaat uit lagen filodeeg, noten en honingsiroop?', 'sv': 'Vilket grekiskt bakverk lager fyllodeg, notter och honungssirap?', 'fr': "Quelle patisserie grecque est composee de couches de pate phyllo, de noix et de sirop de miel ?", 'da': 'Hvilket græsk bagværk er lagdelt med filodej, nødder og honningsirup?'},
 'opt1': {'en': 'Baklava', 'de': 'Baklava', 'es': 'Baklava', 'nl': 'Baklava', 'sv': 'Baklava', 'fr': 'Baklava', 'da': 'Baklava'},
 'opt2': {'en': 'Loukoumades', 'de': 'Loukoumades', 'es': 'Loukoumades', 'nl': 'Loukoumades', 'sv': 'Loukoumades', 'fr': 'Loukoumades', 'da': 'Loukoumades'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which sense organ contains the cornea and retina?', 'de': 'Welches Sinnesorgan enthaelt Hornhaut und Netzhaut?', 'es': '¿Que organo sensorial contiene la cornea y la retina?', 'nl': 'Welk zintuig bevat het hoornvlies en het netvlies?', 'sv': 'Vilket sinnesorgan innehaller hornhinnan och natthinnan?', 'fr': "Quel organe sensoriel contient la cornee et la retine ?", 'da': 'Hvilket sanseorgan indeholder hornhinden og nethinden?'},
 'opt1': {'en': 'Eye', 'de': 'Auge', 'es': 'Ojo', 'nl': 'Oog', 'sv': 'Oga', 'fr': 'Oeil', 'da': 'Øje'},
 'opt2': {'en': 'Ear', 'de': 'Ohr', 'es': 'Oido', 'nl': 'Oor', 'sv': 'Ora', 'fr': 'Oreille', 'da': 'Øre'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ruins of the ancient city of Pompeii?', 'de': 'In welchem Land liegen die Ruinen der antiken Stadt Pompeji?', 'es': '¿En que pais estan las ruinas de la antigua ciudad de Pompeya?', 'nl': 'In welk land liggen de ruines van de oude stad Pompeii?', 'sv': 'I vilket land ligger ruinerna av den antika staden Pompeji?', 'fr': "Dans quel pays se trouvent les ruines de l'ancienne cite de Pompei ?", 'da': 'Hvilket land huser ruinerne af den antikke by Pompeji?'},
 'opt1': {'en': 'Italy', 'de': 'Italien', 'es': 'Italia', 'nl': 'Italie', 'sv': 'Italien', 'fr': "l'Italie", 'da': 'Italien'},
 'opt2': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was a former Miss America contestant advocate before marrying into politics?', 'de': 'Welche First Lady war vor ihrer politischen Ehe eine Fuersprecherin und ehemalige Miss-America-Teilnehmerin?', 'es': '¿Que primera dama fue defensora y ex concursante de Miss America antes de casarse con un politico?', 'nl': 'Welke First Lady was voordat ze met een politicus trouwde een voormalige Miss America-kandidate en pleitbezorger?', 'sv': 'Vilken First Lady var en tidigare Miss America-deltagare och forespråkare innan hon gifte in i politiken?', 'fr': "Quelle Première dame etait une ancienne candidate a Miss America avant d'epouser un homme politique ?", 'da': 'Hvilken førstedame var en tidligere Miss America-deltager, før hun giftede sig ind i politik?'},
 'opt1': {'en': 'Nancy Reagan', 'de': 'Nancy Reagan', 'es': 'Nancy Reagan', 'nl': 'Nancy Reagan', 'sv': 'Nancy Reagan', 'fr': 'Nancy Reagan', 'da': 'Nancy Reagan'},
 'opt2': {'en': 'Betty Ford', 'de': 'Betty Ford', 'es': 'Betty Ford', 'nl': 'Betty Ford', 'sv': 'Betty Ford', 'fr': 'Betty Ford', 'da': 'Betty Ford'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which composer wrote 'Swan Lake' and 'The Nutcracker'?", 'de': "Welcher Komponist schrieb 'Schwanensee' und 'Der Nussknacker'?", 'es': "¿Que compositor escribio 'El lago de los cisnes' y 'El cascanueces'?", 'nl': "Welke componist schreef 'Het Zwanenmeer' en 'De Notenkraker'?", 'sv': "Vilken kompositor skrev 'Svansjon' och 'Nötknäpparen'?", 'fr': "Quel compositeur a ecrit 'Le Lac des cygnes' et 'Casse-Noisette' ?", 'da': "Hvilken komponist skrev 'Svanesøen' og 'Nøddeknækkeren'?"},
 'opt1': {'en': 'Tchaikovsky', 'de': 'Tschaikowski', 'es': 'Chaikovski', 'nl': 'Tsjaikovski', 'sv': 'Tjajkovskij', 'fr': 'Tchaikovski', 'da': 'Tjajkovskij'},
 'opt2': {'en': 'Stravinsky', 'de': 'Strawinsky', 'es': 'Stravinski', 'nl': 'Stravinsky', 'sv': 'Stravinskij', 'fr': 'Stravinski', 'da': 'Stravinskij'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Cher Horowitz in 'Clueless'?", 'de': "Welche Schauspielerin spielte Cher Horowitz in 'Clueless'?", 'es': "¿Que actriz interpreto a Cher Horowitz en 'Fuera de onda'?", 'nl': "Welke actrice speelde Cher Horowitz in 'Clueless'?", 'sv': "Vilken skadespelerska spelade Cher Horowitz i 'Clueless'?", 'fr': "Quelle actrice a joue Cher Horowitz dans 'Clueless' ?", 'da': "Hvilken skuespiller spillede Cher Horowitz i 'Clueless'?"},
 'opt1': {'en': 'Alicia Silverstone', 'de': 'Alicia Silverstone', 'es': 'Alicia Silverstone', 'nl': 'Alicia Silverstone', 'sv': 'Alicia Silverstone', 'fr': 'Alicia Silverstone', 'da': 'Alicia Silverstone'},
 'opt2': {'en': 'Reese Witherspoon', 'de': 'Reese Witherspoon', 'es': 'Reese Witherspoon', 'nl': 'Reese Witherspoon', 'sv': 'Reese Witherspoon', 'fr': 'Reese Witherspoon', 'da': 'Reese Witherspoon'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which fermented soybean paste is used in Japanese miso soup?', 'de': 'Welche fermentierte Sojabohnenpaste wird in japanischer Misosuppe verwendet?', 'es': '¿Que pasta de soja fermentada se usa en la sopa de miso japonesa?', 'nl': 'Welke gefermenteerde sojabonenpasta wordt gebruikt in Japanse misosoep?', 'sv': 'Vilken fermenterad sojabönspasta anvands i japansk misosoppa?', 'fr': "Quelle pate de soja fermentee est utilisee dans la soupe miso japonaise ?", 'da': 'Hvilken fermenteret sojabønnepasta bruges i japansk misosuppe?'},
 'opt1': {'en': 'Miso', 'de': 'Miso', 'es': 'Miso', 'nl': 'Miso', 'sv': 'Miso', 'fr': 'Miso', 'da': 'Miso'},
 'opt2': {'en': 'Tofu', 'de': 'Tofu', 'es': 'Tofu', 'nl': 'Tofu', 'sv': 'Tofu', 'fr': 'Tofu', 'da': 'Tofu'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to Bryce Canyon and Zion National Park?', 'de': 'In welchem US-Bundesstaat liegen der Bryce Canyon und der Zion-Nationalpark?', 'es': '¿En que estado de EE.UU. estan el Cañon Bryce y el Parque Nacional Zion?', 'nl': 'In welke Amerikaanse staat liggen Bryce Canyon en Zion National Park?', 'sv': 'I vilken amerikansk delstat ligger Bryce Canyon och Zion nationalpark?', 'fr': "Dans quel État americain se trouvent Bryce Canyon et le parc national de Zion ?", 'da': 'I hvilken amerikansk delstat ligger Bryce Canyon og Zion nationalpark?'},
 'opt1': {'en': 'Utah', 'de': 'Utah', 'es': 'Utah', 'nl': 'Utah', 'sv': 'Utah', 'fr': "l'Utah", 'da': 'Utah'},
 'opt2': {'en': 'Arizona', 'de': 'Arizona', 'es': 'Arizona', 'nl': 'Arizona', 'sv': 'Arizona', 'fr': "l'Arizona", 'da': 'Arizona'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which insect undergoes complete metamorphosis and is known for silk production in China?', 'de': 'Welches Insekt durchlaeuft eine vollstaendige Metamorphose und ist fuer die Seidenproduktion in China bekannt?', 'es': '¿Que insecto sufre una metamorfosis completa y es conocido por la produccion de seda en China?', 'nl': 'Welk insect ondergaat volledige metamorfose en staat bekend om zijdeproductie in China?', 'sv': 'Vilken insekt genomgar fullstandig metamorfos och ar kand for silkesproduktion i Kina?', 'fr': "Quel insecte subit une metamorphose complete et est connu pour la production de soie en Chine ?", 'da': 'Hvilket insekt gennemgår fuldstændig metamorfose og er kendt for silkeproduktion i Kina?'},
 'opt1': {'en': 'Silkworm moth', 'de': 'Seidenspinner', 'es': 'Gusano de seda', 'nl': 'Zijderupsvlinder', 'sv': 'Silkesfjaril', 'fr': 'Ver a soie', 'da': 'Silkesommerfugl'},
 'opt2': {'en': 'Honeybee', 'de': 'Honigbiene', 'es': 'Abeja melifera', 'nl': 'Honingbij', 'sv': 'Honungsbi', 'fr': 'Abeille', 'da': 'Honningbi'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Chinese festival involves dragon boat races and rice dumplings?', 'de': 'Welches chinesische Fest beinhaltet Drachenbootrennen und Reisknoedel?', 'es': '¿Que festival chino incluye carreras de botes dragon y bolas de arroz?', 'nl': 'Welk Chinees festival omvat drakenbootraces en rijstknoedels?', 'sv': 'Vilken kinesisk festival innebar drakbatstavlingar och risbullar?', 'fr': "Quel festival chinois implique des courses de bateaux-dragons et des boulettes de riz ?", 'da': 'Hvilken kinesisk festival involverer dragebådsløb og risboller?'},
 'opt1': {'en': 'Dragon Boat Festival', 'de': 'Drachenbootfest', 'es': 'Festival del Barco Dragon', 'nl': 'Drakenbootfestival', 'sv': 'Drakbatsfestivalen', 'fr': 'La fete des bateaux-dragons', 'da': 'Dragebådsfestival'},
 'opt2': {'en': 'Mid-Autumn Festival', 'de': 'Mittherbstfest', 'es': 'Festival del Medio Otoño', 'nl': 'Midherfstfestival', 'sv': 'Midhostfestivalen', 'fr': "La fete de la mi-automne", 'da': 'Midtefterårsfestival'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which author wrote 'Rebecca', a gothic romance novel?", 'de': "Welche Autorin schrieb 'Rebecca', einen gotischen Liebesroman?", 'es': "¿Que autora escribio 'Rebeca', una novela romantica gotica?", 'nl': "Welke auteur schreef 'Rebecca', een gotische liefdesroman?", 'sv': "Vilken forfattare skrev 'Rebecca', en gotisk romantisk roman?", 'fr': "Quelle autrice a ecrit 'Rebecca', un roman gothique ?", 'da': "Hvilken forfatter skrev 'Rebecca', en gotisk kærlighedsroman?"},
 'opt1': {'en': 'Daphne du Maurier', 'de': 'Daphne du Maurier', 'es': 'Daphne du Maurier', 'nl': 'Daphne du Maurier', 'sv': 'Daphne du Maurier', 'fr': 'Daphne du Maurier', 'da': 'Daphne du Maurier'},
 'opt2': {'en': 'Agatha Christie', 'de': 'Agatha Christie', 'es': 'Agatha Christie', 'nl': 'Agatha Christie', 'sv': 'Agatha Christie', 'fr': 'Agatha Christie', 'da': 'Agatha Christie'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the Northern Lights viewing town of Tromso?', 'de': 'In welchem Land liegt Tromsø, bekannt fuer Nordlichtbeobachtung?', 'es': '¿En que pais esta Tromso, conocida para ver la aurora boreal?', 'nl': 'In welk land ligt Tromso, bekend om het bekijken van het noorderlicht?', 'sv': 'I vilket land ligger Tromso, kant for norrskensvisning?', 'fr': "Dans quel pays se trouve Tromso, connue pour observer les aurores boreales ?", 'da': 'Hvilket land huser Tromsø, kendt for at se nordlys?'},
 'opt1': {'en': 'Norway', 'de': 'Norwegen', 'es': 'Noruega', 'nl': 'Noorwegen', 'sv': 'Norge', 'fr': 'la Norvège', 'da': 'Norge'},
 'opt2': {'en': 'Iceland', 'de': 'Island', 'es': 'Islandia', 'nl': 'IJsland', 'sv': 'Island', 'fr': "l'Islande", 'da': 'Island'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which instrument is traditionally used in mariachi bands alongside guitars and trumpets?', 'de': 'Welches Instrument wird traditionell in Mariachi-Bands neben Gitarren und Trompeten verwendet?', 'es': '¿Que instrumento se usa tradicionalmente en bandas de mariachi junto con guitarras y trompetas?', 'nl': 'Welk instrument wordt traditioneel gebruikt in mariachi-bands naast gitaren en trompetten?', 'sv': 'Vilket instrument anvands traditionellt i mariachiband tillsammans med gitarrer och trumpeter?', 'fr': "Quel instrument est traditionnellement utilise dans les groupes de mariachis aux cotes des guitares et trompettes ?", 'da': 'Hvilket instrument bruges traditionelt i mariachi-bands sammen med guitarer og trompeter?'},
 'opt1': {'en': 'Vihuela', 'de': 'Vihuela', 'es': 'Vihuela', 'nl': 'Vihuela', 'sv': 'Vihuela', 'fr': 'Vihuela', 'da': 'Vihuela'},
 'opt2': {'en': 'Sitar', 'de': 'Sitar', 'es': 'Sitar', 'nl': 'Sitar', 'sv': 'Sitar', 'fr': 'Sitar', 'da': 'Sitar'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a stew made with okra, often with shrimp or sausage?', 'de': 'Welches suddstaatliche Gericht ist ein Eintopf mit Okra, oft mit Garnelen oder Wurst?', 'es': '¿Que plato sureño de EE.UU. es un guiso hecho con okra, a menudo con camarones o salchicha?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een stoofpot met okra, vaak met garnalen of worst?', 'sv': 'Vilken sydstatlig ratt ar en gryta gjord med okra, ofta med ragor eller korv?', 'fr': "Quel plat du Sud des États-Unis est un ragout de gombo, souvent avec des crevettes ou de la saucisse ?", 'da': 'Hvilken sydstatsret er en gryde lavet med okra, ofte med rejer eller pølse?'},
 'opt1': {'en': 'Gumbo', 'de': 'Gumbo', 'es': 'Gumbo', 'nl': 'Gumbo', 'sv': 'Gumbo', 'fr': 'Gumbo', 'da': 'Gumbo'},
 'opt2': {'en': 'Jambalaya', 'de': 'Jambalaya', 'es': 'Jambalaya', 'nl': 'Jambalaya', 'sv': 'Jambalaya', 'fr': 'Jambalaya', 'da': 'Jambalaya'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which mineral gives blood its red color when bound to oxygen?", 'de': 'Welches Mineral verleiht dem Blut seine rote Farbe, wenn es an Sauerstoff gebunden ist?', 'es': '¿Que mineral le da a la sangre su color rojo cuando se une al oxigeno?', 'nl': 'Welk mineraal geeft bloed zijn rode kleur wanneer het aan zuurstof gebonden is?', 'sv': 'Vilket mineral ger blodet dess roda farg nar det binds till syre?', 'fr': "Quel mineral donne au sang sa couleur rouge lorsqu'il est lie a l'oxygene ?", 'da': 'Hvilket mineral giver blod dets røde farve, når det er bundet til ilt?'},
 'opt1': {'en': 'Iron', 'de': 'Eisen', 'es': 'Hierro', 'nl': 'IJzer', 'sv': 'Jarn', 'fr': 'Fer', 'da': 'Jern'},
 'opt2': {'en': 'Calcium', 'de': 'Kalzium', 'es': 'Calcio', 'nl': 'Calcium', 'sv': 'Kalcium', 'fr': 'Calcium', 'da': 'Calcium'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which film has Sandra Bullock playing a woman who takes in a homeless teen who becomes a football star?', 'de': 'In welchem Film spielt Sandra Bullock eine Frau, die einen obdachlosen Teenager aufnimmt, der zum Football-Star wird?', 'es': '¿En que pelicula Sandra Bullock interpreta a una mujer que acoge a un adolescente sin hogar que se convierte en estrella del futbol americano?', 'nl': 'In welke film speelt Sandra Bullock een vrouw die een dakloze tiener opneemt die een footballster wordt?', 'sv': 'I vilken film spelar Sandra Bullock en kvinna som tar in en hemlos tonaring som blir en fotbollsstjarna?', 'fr': "Dans quel film Sandra Bullock joue-t-elle une femme qui recueille un adolescent sans-abri qui devient une star du football americain ?", 'da': 'I hvilken film spiller Sandra Bullock en kvinde, der tager sig af en hjemløs teenager, der bliver en fodboldstjerne?'},
 'opt1': {'en': 'The Blind Side', 'de': 'Blind Side - Die große Chance', 'es': 'Un Sueño Posible', 'nl': 'The Blind Side', 'sv': 'The Blind Side', 'fr': 'The Blind Side', 'da': 'The Blind Side'},
 'opt2': {'en': 'Miss Congeniality', 'de': 'Miss Undercover', 'es': 'Miss Simpatia', 'nl': 'Miss Congeniality', 'sv': 'Miss Undercover', 'fr': 'Miss Detective', 'da': 'Miss Undercover'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient temple of Borobudur?', 'de': 'In welchem Land liegt der antike Tempel von Borobudur?', 'es': '¿En que pais esta el antiguo templo de Borobudur?', 'nl': 'In welk land ligt de oude tempel van Borobudur?', 'sv': 'I vilket land ligger det antika templet Borobudur?', 'fr': "Dans quel pays se trouve l'ancien temple de Borobudur ?", 'da': 'Hvilket land huser det gamle tempel Borobudur?'},
 'opt1': {'en': 'Indonesia', 'de': 'Indonesien', 'es': 'Indonesia', 'nl': 'Indonesie', 'sv': 'Indonesien', 'fr': "l'Indonesie", 'da': 'Indonesien'},
 'opt2': {'en': 'Vietnam', 'de': 'Vietnam', 'es': 'Vietnam', 'nl': 'Vietnam', 'sv': 'Vietnam', 'fr': 'le Vietnam', 'da': 'Vietnam'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which brand created the iconic Birkin bag?', 'de': 'Welche Marke schuf die ikonische Birkin-Tasche?', 'es': '¿Que marca creo el iconico bolso Birkin?', 'nl': 'Welk merk creëerde de iconische Birkin-tas?', 'sv': 'Vilket marke skapade den ikoniska Birkinvaskan?', 'fr': "Quelle marque a cree l'iconique sac Birkin ?", 'da': 'Hvilket mærke skabte den ikoniske Birkin-taske?'},
 'opt1': {'en': 'Hermes', 'de': 'Hermes', 'es': 'Hermes', 'nl': 'Hermes', 'sv': 'Hermes', 'fr': 'Hermes', 'da': 'Hermes'},
 'opt2': {'en': 'Chanel', 'de': 'Chanel', 'es': 'Chanel', 'nl': 'Chanel', 'sv': 'Chanel', 'fr': 'Chanel', 'da': 'Chanel'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which nocturnal bird is known for its silent flight and hooting call?', 'de': 'Welcher nachtaktive Vogel ist fuer seinen lautlosen Flug und Ruf bekannt?', 'es': '¿Que ave nocturna es conocida por su vuelo silencioso y su canto?', 'nl': 'Welke nachtactieve vogel staat bekend om zijn geluidloze vlucht en roep?', 'sv': 'Vilken nattaktiv fagel ar kand for sin ljudlosa flykt och sitt hoande?', 'fr': "Quel oiseau nocturne est connu pour son vol silencieux et son hululement ?", 'da': 'Hvilken natlig fugl er kendt for sin lydløse flugt og tuden?'},
 'opt1': {'en': 'Owl', 'de': 'Eule', 'es': 'Buho', 'nl': 'Uil', 'sv': 'Uggla', 'fr': 'Hibou', 'da': 'Ugle'},
 'opt2': {'en': 'Bat', 'de': 'Fledermaus', 'es': 'Murcielago', 'nl': 'Vleermuis', 'sv': 'Flader Mus', 'fr': 'Chauve-souris', 'da': 'Flagermus'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was the youngest ever at age 21 when her husband took office?', 'de': 'Welche First Lady war mit 21 Jahren die juengste, als ihr Mann das Amt antrat?', 'es': '¿Que primera dama fue la mas joven, a los 21 años, cuando su esposo asumio el cargo?', 'nl': 'Welke First Lady was op 21-jarige leeftijd de jongste ooit toen haar man aantrad?', 'sv': 'Vilken First Lady var den yngsta nagonsin, 21 ar gammal, nar hennes man tilltradde?', 'fr': "Quelle Première dame etait la plus jeune jamais vue, a 21 ans, lorsque son mari a pris ses fonctions ?", 'da': 'Hvilken førstedame var den yngste nogensinde, 21 år gammel, da hendes mand tiltrådte?'},
 'opt1': {'en': 'Frances Cleveland', 'de': 'Frances Cleveland', 'es': 'Frances Cleveland', 'nl': 'Frances Cleveland', 'sv': 'Frances Cleveland', 'fr': 'Frances Cleveland', 'da': 'Frances Cleveland'},
 'opt2': {'en': 'Julia Tyler', 'de': 'Julia Tyler', 'es': 'Julia Tyler', 'nl': 'Julia Tyler', 'sv': 'Julia Tyler', 'fr': 'Julia Tyler', 'da': 'Julia Tyler'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which First Lady is credited with modernizing White House social events and hosting Camelot-era receptions?', 'de': 'Welche First Lady modernisierte die gesellschaftlichen Ereignisse im Weißen Haus und veranstaltete Empfaenge der Camelot-Ära?', 'es': '¿Que primera dama es reconocida por modernizar los eventos sociales de la Casa Blanca y organizar recepciones de la era Camelot?', 'nl': 'Welke First Lady moderniseerde de sociale evenementen van het Witte Huis en organiseerde Camelot-tijdperk recepties?', 'sv': 'Vilken First Lady moderniserade Vita husets sociala evenemang och stod vard for Camelot-eran mottagningar?', 'fr': "Quelle Première dame a modernise les evenements mondains de la Maison-Blanche et organise des receptions de l'ere Camelot ?", 'da': 'Hvilken førstedame moderniserede Det Hvide Hus’ sociale begivenheder og var vært for Camelot-æraens receptioner?'},
 'opt1': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'},
 'opt2': {'en': 'Mamie Eisenhower', 'de': 'Mamie Eisenhower', 'es': 'Mamie Eisenhower', 'nl': 'Mamie Eisenhower', 'sv': 'Mamie Eisenhower', 'fr': 'Mamie Eisenhower', 'da': 'Mamie Eisenhower'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which British author wrote 'Mrs Dalloway'?", 'de': "Welche britische Autorin schrieb 'Mrs Dalloway'?", 'es': "¿Que autora britanica escribio 'La señora Dalloway'?", 'nl': "Welke Britse auteur schreef 'Mrs Dalloway'?", 'sv': "Vilken brittisk forfattare skrev 'Mrs Dalloway'?", 'fr': "Quelle autrice britannique a ecrit 'Mrs Dalloway' ?", 'da': "Hvilken britisk forfatter skrev 'Fru Dalloway'?"},
 'opt1': {'en': 'Virginia Woolf', 'de': 'Virginia Woolf', 'es': 'Virginia Woolf', 'nl': 'Virginia Woolf', 'sv': 'Virginia Woolf', 'fr': 'Virginia Woolf', 'da': 'Virginia Woolf'},
 'opt2': {'en': 'Agatha Christie', 'de': 'Agatha Christie', 'es': 'Agatha Christie', 'nl': 'Agatha Christie', 'sv': 'Agatha Christie', 'fr': 'Agatha Christie', 'da': 'Agatha Christie'}},

{'category': 'Science', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which vitamin deficiency causes rickets, a bone-softening disease in children?', 'de': 'Welcher Vitaminmangel verursacht Rachitis, eine knochenerweichende Krankheit bei Kindern?', 'es': '¿Que deficiencia vitaminica causa el raquitismo, una enfermedad que ablanda los huesos en niños?', 'nl': 'Welk vitaminetekort veroorzaakt rachitis, een botverwekende ziekte bij kinderen?', 'sv': 'Vilken vitaminbrist orsakar rakitis, en bensjukdom hos barn?', 'fr': "Quelle carence en vitamine cause le rachitisme, une maladie qui ramollit les os chez les enfants ?", 'da': 'Hvilken vitaminmangel forårsager rakitis, en knogleblødgørende sygdom hos børn?'},
 'opt1': {'en': 'Vitamin D', 'de': 'Vitamin D', 'es': 'Vitamina D', 'nl': 'Vitamine D', 'sv': 'Vitamin D', 'fr': 'Vitamine D', 'da': 'Vitamin D'},
 'opt2': {'en': 'Vitamin K', 'de': 'Vitamin K', 'es': 'Vitamina K', 'nl': 'Vitamine K', 'sv': 'Vitamin K', 'fr': 'Vitamine K', 'da': 'Vitamin K'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Norwegian painter created 'The Scream'?", 'de': "Welcher norwegische Maler schuf 'Der Schrei'?", 'es': "¿Que pintor noruego creo 'El grito'?", 'nl': "Welke Noorse schilder creëerde 'De Schreeuw'?", 'sv': "Vilken norsk malare skapade 'Skriet'?", 'fr': "Quel peintre norvegien a cree 'Le Cri' ?", 'da': "Hvilken norsk maler skabte 'Skriget'?"},
 'opt1': {'en': 'Edvard Munch', 'de': 'Edvard Munch', 'es': 'Edvard Munch', 'nl': 'Edvard Munch', 'sv': 'Edvard Munch', 'fr': 'Edvard Munch', 'da': 'Edvard Munch'},
 'opt2': {'en': 'Gustav Klimt', 'de': 'Gustav Klimt', 'es': 'Gustav Klimt', 'nl': 'Gustav Klimt', 'sv': 'Gustav Klimt', 'fr': 'Gustav Klimt', 'da': 'Gustav Klimt'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Central Asian country is famous for its nomadic yurts and vast steppes?', 'de': 'Welches zentralasiatische Land ist beruehmt fuer seine nomadischen Jurten und weiten Steppen?', 'es': '¿Que pais de Asia Central es famoso por sus yurtas nomadas y vastas estepas?', 'nl': 'Welk Centraal-Aziatisch land is beroemd om zijn nomadische jurten en uitgestrekte steppen?', 'sv': 'Vilket centralasiatiskt land ar kant for sina nomadiska jurtor och vidstrackta stapper?', 'fr': "Quel pays d'Asie centrale est celebre pour ses yourtes nomades et ses vastes steppes ?", 'da': 'Hvilket centralasiatisk land er berømt for sine nomadiske jurter og vidtstrakte stepper?'},
 'opt1': {'en': 'Mongolia', 'de': 'Mongolei', 'es': 'Mongolia', 'nl': 'Mongolie', 'sv': 'Mongoliet', 'fr': 'la Mongolie', 'da': 'Mongoliet'},
 'opt2': {'en': 'Kazakhstan', 'de': 'Kasachstan', 'es': 'Kazajistan', 'nl': 'Kazachstan', 'sv': 'Kazakstan', 'fr': 'le Kazakhstan', 'da': 'Kasakhstan'}},

{'category': 'Music', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Italian composer wrote the opera 'La Traviata'?", 'de': "Welcher italienische Komponist schrieb die Oper 'La Traviata'?", 'es': "¿Que compositor italiano escribio la opera 'La Traviata'?", 'nl': "Welke Italiaanse componist schreef de opera 'La Traviata'?", 'sv': "Vilken italiensk kompositor skrev operan 'La Traviata'?", 'fr': "Quel compositeur italien a ecrit l'opera 'La Traviata' ?", 'da': "Hvilken italiensk komponist skrev operaen 'La Traviata'?"},
 'opt1': {'en': 'Giuseppe Verdi', 'de': 'Giuseppe Verdi', 'es': 'Giuseppe Verdi', 'nl': 'Giuseppe Verdi', 'sv': 'Giuseppe Verdi', 'fr': 'Giuseppe Verdi', 'da': 'Giuseppe Verdi'},
 'opt2': {'en': 'Giacomo Puccini', 'de': 'Giacomo Puccini', 'es': 'Giacomo Puccini', 'nl': 'Giacomo Puccini', 'sv': 'Giacomo Puccini', 'fr': 'Giacomo Puccini', 'da': 'Giacomo Puccini'}},

{'category': 'Food', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which French dish is a savory custard tart, often with bacon and cheese?', 'de': 'Welches franzoesische Gericht ist eine herzhafte Eiertorte, oft mit Speck und Kaese?', 'es': '¿Que plato frances es una tarta de natillas salada, a menudo con tocino y queso?', 'nl': 'Welk Frans gerecht is een hartige eiertaart, vaak met spek en kaas?', 'sv': 'Vilken fransk ratt ar en salt äggkraskaka, ofta med bacon och ost?', 'fr': "Quel plat français est une tarte a base de flan sale, souvent au lard et au fromage ?", 'da': 'Hvilken fransk ret er en salt æggetærte, ofte med bacon og ost?'},
 'opt1': {'en': 'Quiche', 'de': 'Quiche', 'es': 'Quiche', 'nl': 'Quiche', 'sv': 'Quiche', 'fr': 'Quiche', 'da': 'Quiche'},
 'opt2': {'en': 'Souffle', 'de': 'Souffle', 'es': 'Souffle', 'nl': 'Souffle', 'sv': 'Souffle', 'fr': 'Souffle', 'da': 'Souffle'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Roman mythology, which god is the counterpart of the Greek god Zeus?', 'de': 'Welcher roemische Gott ist das Gegenstueck zum griechischen Gott Zeus?', 'es': '¿Que dios romano es la contraparte del dios griego Zeus?', 'nl': 'Welke Romeinse god is de tegenhanger van de Griekse god Zeus?', 'sv': 'Vilken romersk gud ar motsvarigheten till den grekiska guden Zeus?', 'fr': "Quel dieu romain est l'equivalent du dieu grec Zeus ?", 'da': 'Hvilken romersk gud er modstykket til den græske gud Zeus?'},
 'opt1': {'en': 'Jupiter', 'de': 'Jupiter', 'es': 'Jupiter', 'nl': 'Jupiter', 'sv': 'Jupiter', 'fr': 'Jupiter', 'da': 'Jupiter'},
 'opt2': {'en': 'Neptune', 'de': 'Neptun', 'es': 'Neptuno', 'nl': 'Neptunus', 'sv': 'Neptunus', 'fr': 'Neptune', 'da': 'Neptun'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which actress won an Oscar for playing Queen Elizabeth II in 'The Queen'?", 'de': "Welche Schauspielerin gewann einen Oscar fuer ihre Rolle als Koenigin Elisabeth II. in 'The Queen'?", 'es': "¿Que actriz gano un Oscar por interpretar a la Reina Isabel II en 'The Queen'?", 'nl': "Welke actrice won een Oscar voor haar rol als Koningin Elizabeth II in 'The Queen'?", 'sv': "Vilken skadespelerska vann en Oscar for att ha spelat drottning Elizabeth II i 'The Queen'?", 'fr': "Quelle actrice a remporte un Oscar pour son role de la reine Elisabeth II dans 'The Queen' ?", 'da': "Hvilken skuespiller vandt en Oscar for at spille Dronning Elizabeth II i 'The Queen'?"},
 'opt1': {'en': 'Helen Mirren', 'de': 'Helen Mirren', 'es': 'Helen Mirren', 'nl': 'Helen Mirren', 'sv': 'Helen Mirren', 'fr': 'Helen Mirren', 'da': 'Helen Mirren'},
 'opt2': {'en': 'Judi Dench', 'de': 'Judi Dench', 'es': 'Judi Dench', 'nl': 'Judi Dench', 'sv': 'Judi Dench', 'fr': 'Judi Dench', 'da': 'Judi Dench'}},

{'category': 'Space', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which planet has visible rings made mostly of ice and rock, tilted at a dramatic angle?', 'de': 'Welcher Planet hat sichtbare Ringe aus Eis und Gestein, in einem dramatischen Winkel geneigt?', 'es': '¿Que planeta tiene anillos visibles hechos principalmente de hielo y roca, inclinados en un angulo dramatico?', 'nl': 'Welke planeet heeft zichtbare ringen van voornamelijk ijs en gesteente, gekanteld in een dramatische hoek?', 'sv': 'Vilken planet har synliga ringar gjorda mestadels av is och sten, lutade i en dramatisk vinkel?', 'fr': "Quelle planete a des anneaux visibles faits principalement de glace et de roche, inclines a un angle spectaculaire ?", 'da': 'Hvilken planet har synlige ringe lavet mest af is og sten, vippet i en dramatisk vinkel?'},
 'opt1': {'en': 'Saturn', 'de': 'Saturn', 'es': 'Saturno', 'nl': 'Saturnus', 'sv': 'Saturnus', 'fr': 'Saturne', 'da': 'Saturn'},
 'opt2': {'en': 'Uranus', 'de': 'Uranus', 'es': 'Urano', 'nl': 'Uranus', 'sv': 'Uranus', 'fr': 'Uranus', 'da': 'Uranus'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which country's flag features a yin-yang-like symbol surrounded by four trigrams?", 'de': 'Welches Land hat auf seiner Flagge ein Yin-Yang-aehnliches Symbol umgeben von vier Trigrammen?', 'es': '¿Que pais tiene en su bandera un simbolo parecido al yin-yang rodeado de cuatro trigramas?', 'nl': 'Welk land heeft op zijn vlag een yin-yang-achtig symbool omringd door vier trigrammen?', 'sv': 'Vilket lands flagga har en yin-yang-liknande symbol omgiven av fyra trigram?', 'fr': "Le drapeau de quel pays presente un symbole semblable au yin-yang entoure de quatre trigrammes ?", 'da': 'Hvilket lands flag har et yin-yang-lignende symbol omgivet af fire trigrammer?'},
 'opt1': {'en': 'South Korea', 'de': 'Suedkorea', 'es': 'Corea del Sur', 'nl': 'Zuid-Korea', 'sv': 'Sydkorea', 'fr': 'la Coree du Sud', 'da': 'Sydkorea'},
 'opt2': {'en': 'Japan', 'de': 'Japan', 'es': 'Japon', 'nl': 'Japan', 'sv': 'Japan', 'fr': 'le Japon', 'da': 'Japan'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which French couturier is credited with liberating women from corsets in the early 1900s?', 'de': 'Welcher franzoesische Couturier befreite Frauen zu Beginn des 20. Jahrhunderts von Korsetts?', 'es': '¿Que couturier frances es reconocido por liberar a las mujeres de los corses a principios del siglo XX?', 'nl': 'Welke Franse couturier bevrijdde vrouwen begin 1900 van korsetten?', 'sv': 'Vilken fransk modeskapare tillskrivs att ha befriat kvinnor fran korsetter i borjan av 1900-talet?', 'fr': "Quel couturier français est credite d'avoir libere les femmes du corset au debut des annees 1900 ?", 'da': 'Hvilken fransk couturier krediteres for at have befriet kvinder fra korsetter i begyndelsen af 1900-tallet?'},
 'opt1': {'en': 'Paul Poiret', 'de': 'Paul Poiret', 'es': 'Paul Poiret', 'nl': 'Paul Poiret', 'sv': 'Paul Poiret', 'fr': 'Paul Poiret', 'da': 'Paul Poiret'},
 'opt2': {'en': 'Charles Worth', 'de': 'Charles Worth', 'es': 'Charles Worth', 'nl': 'Charles Worth', 'sv': 'Charles Worth', 'fr': 'Charles Worth', 'da': 'Charles Worth'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which ancient Greek philosopher is said to have lived in a barrel and carried a lamp searching for an honest man?', 'de': 'Welcher antike griechische Philosoph soll in einem Fass gelebt und mit einer Lampe nach einem ehrlichen Mann gesucht haben?', 'es': '¿Que filosofo griego antiguo se dice que vivio en un barril y llevaba una lampara buscando a un hombre honesto?', 'nl': 'Welke oude Griekse filosoof zou in een vat hebben geleefd en met een lamp op zoek zijn gegaan naar een eerlijk man?', 'sv': 'Vilken antik grekisk filosof sags ha bott i en tunna och burit en lampa pa jakt efter en arlig man?', 'fr': "Quel philosophe grec antique aurait vecu dans un tonneau et porte une lampe a la recherche d'un homme honnete ?", 'da': 'Hvilken oldgræsk filosof siges at have levet i en tønde og båret en lampe på jagt efter en ærlig mand?'},
 'opt1': {'en': 'Diogenes', 'de': 'Diogenes', 'es': 'Diogenes', 'nl': 'Diogenes', 'sv': 'Diogenes', 'fr': 'Diogene', 'da': 'Diogenes'},
 'opt2': {'en': 'Aristotle', 'de': 'Aristoteles', 'es': 'Aristoteles', 'nl': 'Aristoteles', 'sv': 'Aristoteles', 'fr': 'Aristote', 'da': 'Aristoteles'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which quilting pattern resembles overlapping fans and was popular in Victorian crazy quilts?', 'de': 'Welches Quiltmuster ahnelt uberlappenden Faechern und war bei viktorianischen Crazy Quilts beliebt?', 'es': '¿Que patron de acolchado se asemeja a abanicos superpuestos y era popular en los edredones victorianos crazy quilt?', 'nl': 'Welk quiltpatroon lijkt op overlappende waaiers en was populair in Victoriaanse crazy quilts?', 'sv': 'Vilket lapptacksmonster liknar overlappande solfjadrar och var populart i viktorianska crazy quilts?', 'fr': "Quel motif de patchwork ressemble a des eventails superposes et etait populaire dans les crazy quilts victoriens ?", 'da': 'Hvilket quiltmønster ligner overlappende vifter og var populært i victorianske crazy quilts?'},
 'opt1': {'en': 'Fan pattern', 'de': 'Faechermuster', 'es': 'Patron de abanico', 'nl': 'Waaierpatroon', 'sv': 'Solfjadersmonster', 'fr': "Motif d'eventail", 'da': 'Vifte-mønster'},
 'opt2': {'en': 'Wedding Ring pattern', 'de': 'Ehering-Muster', 'es': 'Patron de anillo de boda', 'nl': 'Trouwringpatroon', 'sv': 'Vigselringsmonster', 'fr': 'Motif anneau de mariage', 'da': 'Vielsesring-mønster'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which desert plant stores water in its thick stems and is often mistaken for a cactus?', 'de': 'Welche Wuestenpflanze speichert Wasser in ihren dicken Stengeln und wird oft mit einem Kaktus verwechselt?', 'es': '¿Que planta del desierto almacena agua en sus tallos gruesos y a menudo se confunde con un cactus?', 'nl': 'Welke woestijnplant slaat water op in zijn dikke stengels en wordt vaak verward met een cactus?', 'sv': 'Vilken okenvaxt lagrar vatten i sina tjocka stjalkar och forvaxlas ofta med en kaktus?', 'fr': "Quelle plante du desert stocke de l'eau dans ses tiges epaisses et est souvent confondue avec un cactus ?", 'da': 'Hvilken ørkenplante opbevarer vand i sine tykke stængler og forveksles ofte med en kaktus?'},
 'opt1': {'en': 'Euphorbia', 'de': 'Euphorbia', 'es': 'Euforbia', 'nl': 'Euphorbia', 'sv': 'Törel (Euphorbia)', 'fr': 'Euphorbe', 'da': 'Vortemælk (Euphorbia)'},
 'opt2': {'en': 'Aloe vera', 'de': 'Aloe vera', 'es': 'Aloe vera', 'nl': 'Aloë vera', 'sv': 'Aloe vera', 'fr': 'Aloe vera', 'da': 'Aloe vera'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which architect designed New York's Guggenheim Museum with its spiral ramp?", 'de': 'Welcher Architekt entwarf das Guggenheim-Museum in New York mit seiner Spiralrampe?', 'es': '¿Que arquitecto diseño el Museo Guggenheim de Nueva York con su rampa en espiral?', 'nl': 'Welke architect ontwierp het Guggenheim Museum in New York met zijn spiraalvormige helling?', 'sv': 'Vilken arkitekt designade Guggenheimmuseet i New York med sin spiralformade ramp?', 'fr': "Quel architecte a conçu le musee Guggenheim de New York avec sa rampe en spirale ?", 'da': 'Hvilken arkitekt designede New Yorks Guggenheim-museum med dets spiralrampe?'},
 'opt1': {'en': 'Frank Lloyd Wright', 'de': 'Frank Lloyd Wright', 'es': 'Frank Lloyd Wright', 'nl': 'Frank Lloyd Wright', 'sv': 'Frank Lloyd Wright', 'fr': 'Frank Lloyd Wright', 'da': 'Frank Lloyd Wright'},
 'opt2': {'en': 'Frank Gehry', 'de': 'Frank Gehry', 'es': 'Frank Gehry', 'nl': 'Frank Gehry', 'sv': 'Frank Gehry', 'fr': 'Frank Gehry', 'da': 'Frank Gehry'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Roman mythology, which goddess is the counterpart of the Greek Aphrodite?', 'de': 'Welche roemische Goettin ist das Gegenstueck zur griechischen Aphrodite?', 'es': '¿Que diosa romana es la contraparte de la griega Afrodita?', 'nl': 'Welke Romeinse godin is de tegenhanger van de Griekse Afrodite?', 'sv': 'Vilken romersk gudinna ar motsvarigheten till den grekiska Afrodite?', 'fr': "Quelle deesse romaine est l'equivalent de l'Aphrodite grecque ?", 'da': 'Hvilken romersk gudinde er modstykket til den græske Afrodite?'},
 'opt1': {'en': 'Venus', 'de': 'Venus', 'es': 'Venus', 'nl': 'Venus', 'sv': 'Venus', 'fr': 'Venus', 'da': 'Venus'},
 'opt2': {'en': 'Diana', 'de': 'Diana', 'es': 'Diana', 'nl': 'Diana', 'sv': 'Diana', 'fr': 'Diane', 'da': 'Diana'}},
]
