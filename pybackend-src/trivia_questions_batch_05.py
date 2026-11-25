# -*- coding: utf-8 -*-
# Trivia Challenge batch #5. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which white grain is a staple food in many Asian countries?', 'de': 'Welches weisse Getreide ist ein Grundnahrungsmittel in vielen asiatischen Laendern?', 'es': '¿Que grano blanco es un alimento basico en muchos paises asiaticos?', 'nl': 'Welke witte graansoort is een basisvoedsel in veel Aziatische landen?', 'sv': 'Vilket vitt sadesslag ar en basfoda i manga asiatiska lander?', 'fr': "Quelle cereale blanche est un aliment de base dans de nombreux pays asiatiques ?", 'da': 'Hvilken hvid kerne er en basisfødevare i mange asiatiske lande?'},
 'opt1': {'en': 'Rice', 'de': 'Reis', 'es': 'Arroz', 'nl': 'Rijst', 'sv': 'Ris', 'fr': 'Le riz', 'da': 'Ris'},
 'opt2': {'en': 'Wheat', 'de': 'Weizen', 'es': 'Trigo', 'nl': 'Tarwe', 'sv': 'Vete', 'fr': 'Le ble', 'da': 'Hvede'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Athens?', 'de': 'In welchem Land liegt die antike Stadt Athen?', 'es': '¿En que pais esta la antigua ciudad de Atenas?', 'nl': 'In welk land ligt de oude stad Athene?', 'sv': 'I vilket land ligger den antika staden Aten?', 'fr': "Dans quel pays se trouve l'ancienne cite d'Athenes ?", 'da': 'Hvilket land huser den antikke by Athen?'},
 'opt1': {'en': 'Turkey', 'de': 'Tuerkei', 'es': 'Turquia', 'nl': 'Turkije', 'sv': 'Turkiet', 'fr': 'la Turquie', 'da': 'Tyrkiet'},
 'opt2': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a boy who can see dead people?', 'de': 'In welchem Filmklassiker kann ein Junge Tote sehen?', 'es': '¿Que pelicula clasica presenta a un niño que puede ver a los muertos?', 'nl': 'In welke klassieke film kan een jongen dode mensen zien?', 'sv': 'I vilken klassisk film kan en pojke se doda manniskor?', 'fr': "Quel film classique met en scene un garçon qui peut voir des morts ?", 'da': 'Hvilken klassisk film har en dreng, der kan se døde mennesker?'},
 'opt1': {'en': 'The Sixth Sense', 'de': 'The Sixth Sense', 'es': 'El sexto sentido', 'nl': 'The Sixth Sense', 'sv': 'Sjätte sinnet', 'fr': 'Sixieme sens', 'da': 'The Sixth Sense'},
 'opt2': {'en': 'Ghost', 'de': 'Ghost - Nachricht von Sam', 'es': 'Ghost', 'nl': 'Ghost', 'sv': 'Ghost', 'fr': 'Ghost', 'da': 'Ghost'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which large marine mammal is known for its song and is the largest animal on Earth?', 'de': 'Welches grosse Meeressaeugetier ist fuer seinen Gesang bekannt und das groesste Tier der Erde?', 'es': '¿Que gran mamifero marino es conocido por su canto y es el animal mas grande de la Tierra?', 'nl': 'Welk groot zeezoogdier staat bekend om zijn gezang en is het grootste dier op aarde?', 'sv': 'Vilket stort havslevande daggdjur ar kant for sin sang och ar jordens storsta djur?', 'fr': "Quel grand mammifere marin est connu pour son chant et est le plus grand animal de la Terre ?", 'da': 'Hvilket stort havpattedyr er kendt for sin sang og er det største dyr på Jorden?'},
 'opt1': {'en': 'Humpback whale', 'de': 'Buckelwal', 'es': 'Ballena jorobada', 'nl': 'Bultrug', 'sv': 'Knolval', 'fr': 'Baleine a bosse', 'da': 'Pukkelhval'},
 'opt2': {'en': 'Blue whale', 'de': 'Blauwal', 'es': 'Ballena azul', 'nl': 'Blauwe vinvis', 'sv': 'Blavalen', 'fr': 'Baleine bleue', 'da': 'Blåhval'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which brown spread is made from roasted cocoa beans and sugar?', 'de': 'Welcher braune Aufstrich wird aus gerösteten Kakaobohnen und Zucker hergestellt?', 'es': '¿Que untable marron se hace con granos de cacao tostados y azucar?', 'nl': 'Welke bruine smeersel wordt gemaakt van gebrande cacaobonen en suiker?', 'sv': 'Vilket brunt palagg gors av rostade kakaobonor och socker?', 'fr': "Quelle tartinade brune est faite de feves de cacao torrefiees et de sucre ?", 'da': 'Hvilket brunt pålæg er lavet af ristede kakaobønner og sukker?'},
 'opt1': {'en': 'Chocolate', 'de': 'Schokolade', 'es': 'Chocolate', 'nl': 'Chocolade', 'sv': 'Choklad', 'fr': 'Chocolat', 'da': 'Chokolade'},
 'opt2': {'en': 'Caramel', 'de': 'Karamell', 'es': 'Caramelo', 'nl': 'Karamel', 'sv': 'Karamell', 'fr': 'Caramel', 'da': 'Karamel'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the Leaning Tower of Pisa?', 'de': 'In welchem Land liegt der Schiefe Turm von Pisa?', 'es': '¿En que pais esta la Torre Inclinada de Pisa?', 'nl': 'In welk land ligt de Scheve Toren van Pisa?', 'sv': 'I vilket land ligger det lutande tornet i Pisa?', 'fr': "Dans quel pays se trouve la tour penchee de Pise ?", 'da': 'Hvilket land huser det skæve tårn i Pisa?'},
 'opt1': {'en': 'Spain', 'de': 'Spanien', 'es': 'España', 'nl': 'Spanje', 'sv': 'Spanien', 'fr': "l'Espagne", 'da': 'Spanien'},
 'opt2': {'en': 'Italy', 'de': 'Italien', 'es': 'Italia', 'nl': 'Italie', 'sv': 'Italien', 'fr': "l'Italie", 'da': 'Italien'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. president served two non-consecutive terms?', 'de': 'Welcher US-Praesident diente zwei nicht aufeinanderfolgende Amtszeiten?', 'es': '¿Que presidente de EE.UU. sirvio dos mandatos no consecutivos?', 'nl': 'Welke Amerikaanse president diende twee niet-opeenvolgende termijnen?', 'sv': 'Vilken amerikansk president tjanstgjorde tva icke-pafoljande mandatperioder?', 'fr': "Quel president americain a exerce deux mandats non consecutifs ?", 'da': 'Hvilken amerikansk præsident sad to ikke-sammenhængende embedsperioder?'},
 'opt1': {'en': 'Grover Cleveland', 'de': 'Grover Cleveland', 'es': 'Grover Cleveland', 'nl': 'Grover Cleveland', 'sv': 'Grover Cleveland', 'fr': 'Grover Cleveland', 'da': 'Grover Cleveland'},
 'opt2': {'en': 'William Taft', 'de': 'William Taft', 'es': 'William Taft', 'nl': 'William Taft', 'sv': 'William Taft', 'fr': 'William Taft', 'da': 'William Taft'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which art form uses clay to create pottery on a spinning wheel?', 'de': 'Welche Kunstform verwendet Ton, um Toepferwaren auf einer Drehscheibe zu erstellen?', 'es': '¿Que forma de arte usa arcilla para crear ceramica en un torno giratorio?', 'nl': 'Welke kunstvorm gebruikt klei om aardewerk te maken op een draaischijf?', 'sv': 'Vilken konstform anvander lera for att skapa keramik pa en snurrande skiva?', 'fr': "Quelle forme d'art utilise de l'argile pour creer de la poterie sur un tour ?", 'da': 'Hvilken kunstform bruger ler til at skabe keramik på en drejeskive?'},
 'opt1': {'en': 'Sculpture', 'de': 'Bildhauerei', 'es': 'Escultura', 'nl': 'Beeldhouwkunst', 'sv': 'Skulptur', 'fr': 'La sculpture', 'da': 'Skulptur'},
 'opt2': {'en': 'Pottery', 'de': 'Toepferei', 'es': 'Alfareria', 'nl': 'Pottenbakken', 'sv': 'Keramik', 'fr': 'La poterie', 'da': 'Keramik'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which instrument has 88 keys and is played with both hands?', 'de': 'Welches Instrument hat 88 Tasten und wird mit beiden Haenden gespielt?', 'es': '¿Que instrumento tiene 88 teclas y se toca con ambas manos?', 'nl': 'Welk instrument heeft 88 toetsen en wordt met beide handen bespeeld?', 'sv': 'Vilket instrument har 88 tangenter och spelas med bada handerna?', 'fr': "Quel instrument possede 88 touches et se joue avec les deux mains ?", 'da': 'Hvilket instrument har 88 tangenter og spilles med begge hænder?'},
 'opt1': {'en': 'Piano', 'de': 'Klavier', 'es': 'Piano', 'nl': 'Piano', 'sv': 'Piano', 'fr': 'Piano', 'da': 'Klaver'},
 'opt2': {'en': 'Organ', 'de': 'Orgel', 'es': 'Organo', 'nl': 'Orgel', 'sv': 'Orgel', 'fr': 'Orgue', 'da': 'Orgel'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which season is associated with falling leaves and cooler weather?', 'de': 'Welche Jahreszeit wird mit fallenden Blaettern und kuehlerem Wetter assoziiert?', 'es': '¿Que estacion se asocia con la caida de hojas y clima mas fresco?', 'nl': 'Welk seizoen wordt geassocieerd met vallende bladeren en koeler weer?', 'sv': 'Vilken arstid forknippas med fallande lov och svalare vader?', 'fr': "Quelle saison est associee aux feuilles qui tombent et au temps plus frais ?", 'da': 'Hvilken årstid forbindes med faldende blade og køligere vejr?'},
 'opt1': {'en': 'Autumn', 'de': 'Herbst', 'es': 'Otoño', 'nl': 'Herfst', 'sv': 'Host', 'fr': 'Automne', 'da': 'Efterår'},
 'opt2': {'en': 'Spring', 'de': 'Fruehling', 'es': 'Primavera', 'nl': 'Lente', 'sv': 'Var', 'fr': 'Printemps', 'da': 'Forår'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which classic film features a young lion who becomes king after his father's death?", 'de': 'In welchem Filmklassiker wird ein junger Loewe nach dem Tod seines Vaters Koenig?', 'es': '¿Que pelicula clasica presenta a un joven leon que se convierte en rey tras la muerte de su padre?', 'nl': 'In welke klassieke film wordt een jonge leeuw koning na de dood van zijn vader?', 'sv': 'I vilken klassisk film blir ett ungt lejon kung efter sin faders dod?', 'fr': "Quel film classique met en scene un jeune lion qui devient roi apres la mort de son pere ?", 'da': 'Hvilken klassisk film har en ung løve, der bliver konge efter sin fars død?'},
 'opt1': {'en': 'The Lion King', 'de': 'Der Koenig der Loewen', 'es': 'El Rey Leon', 'nl': 'The Lion King', 'sv': 'Lejonkungen', 'fr': 'Le Roi Lion', 'da': 'Løvernes Konge'},
 'opt2': {'en': 'Madagascar', 'de': 'Madagascar', 'es': 'Madagascar', 'nl': 'Madagascar', 'sv': 'Madagaskar', 'fr': 'Madagascar', 'da': 'Madagascar'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which U.S. state is known as the 'Show Me State'?", 'de': "Welcher US-Bundesstaat ist als 'Show Me State' bekannt?", 'es': "¿Que estado de EE.UU. es conocido como el 'Show Me State'?", 'nl': "Welke Amerikaanse staat staat bekend als de 'Show Me State'?", 'sv': "Vilken amerikansk delstat ar kand som 'Show Me State'?", 'fr': "Quel État americain est connu comme le 'Show Me State' ?", 'da': "Hvilken amerikansk delstat er kendt som 'Show Me State'?"},
 'opt1': {'en': 'Kansas', 'de': 'Kansas', 'es': 'Kansas', 'nl': 'Kansas', 'sv': 'Kansas', 'fr': 'le Kansas', 'da': 'Kansas'},
 'opt2': {'en': 'Missouri', 'de': 'Missouri', 'es': 'Misuri', 'nl': 'Missouri', 'sv': 'Missouri', 'fr': 'le Missouri', 'da': 'Missouri'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which green vegetable is a key ingredient in traditional pesto?', 'de': 'Welches gruene Gemuese ist eine Schluesselzutat im traditionellen Pesto?', 'es': '¿Que verdura verde es un ingrediente clave en el pesto tradicional?', 'nl': 'Welke groene groente is een sleutelingredient in traditionele pesto?', 'sv': 'Vilken gron gronsak ar en nyckelingrediens i traditionell pesto?', 'fr': "Quel legume vert est un ingredient cle du pesto traditionnel ?", 'da': 'Hvilken grøn grøntsag er en nøgleingrediens i traditionel pesto?'},
 'opt1': {'en': 'Spinach', 'de': 'Spinat', 'es': 'Espinaca', 'nl': 'Spinazie', 'sv': 'Spenat', 'fr': 'Les epinards', 'da': 'Spinat'},
 'opt2': {'en': 'Basil', 'de': 'Basilikum', 'es': 'Albahaca', 'nl': 'Basilicum', 'sv': 'Basilika', 'fr': 'Le basilic', 'da': 'Basilikum'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves hitting a ball with a bat and running around bases?', 'de': 'Bei welcher Sportart schlaegt man einen Ball mit einem Schlaeger und laeuft um Bassen?', 'es': '¿Que deporte consiste en golpear una pelota con un bate y correr alrededor de bases?', 'nl': 'Bij welke sport sla je een bal met een knuppel en ren je rond honken?', 'sv': 'Vilken sport innebar att slå en boll med ett slagtra och springa runt baser?', 'fr': "Quel sport consiste a frapper une balle avec une batte et a courir autour de bases ?", 'da': 'Hvilken sport involverer at slå en bold med et bat og løbe rundt om baser?'},
 'opt1': {'en': 'Baseball', 'de': 'Baseball', 'es': 'Beisbol', 'nl': 'Honkbal', 'sv': 'Baseboll', 'fr': 'Le baseball', 'da': 'Baseball'},
 'opt2': {'en': 'Cricket', 'de': 'Cricket', 'es': 'Criquet', 'nl': 'Cricket', 'sv': 'Cricket', 'fr': 'Le cricket', 'da': 'Cricket'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which small pet is known for hopping and living in a hutch?', 'de': 'Welches kleine Haustier ist fuer sein Huepfen bekannt und lebt in einem Stall?', 'es': '¿Que pequeña mascota es conocida por saltar y vivir en una conejera?', 'nl': 'Welk klein huisdier staat bekend om huppelen en leeft in een hok?', 'sv': 'Vilket litet sallskapsdjur ar kant for att hoppa och bor i en bur?', 'fr': "Quel petit animal de compagnie est connu pour bondir et vivre dans un clapier ?", 'da': 'Hvilket lille kæledyr er kendt for at hoppe og bo i en hytte?'},
 'opt1': {'en': 'Guinea pig', 'de': 'Meerschweinchen', 'es': 'Cobaya', 'nl': 'Cavia', 'sv': 'Marsvin', 'fr': "Cochon d'Inde", 'da': 'Marsvin'},
 'opt2': {'en': 'Rabbit', 'de': 'Kaninchen', 'es': 'Conejo', 'nl': 'Konijn', 'sv': 'Kanin', 'fr': 'Lapin', 'da': 'Kanin'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which holiday celebrates the birth of Jesus Christ in December?', 'de': 'Welcher Feiertag im Dezember feiert die Geburt Jesu Christi?', 'es': '¿Que festividad de diciembre celebra el nacimiento de Jesucristo?', 'nl': 'Welke feestdag in december viert de geboorte van Jezus Christus?', 'sv': 'Vilken hogtid i december firar Jesu Kristi fodelse?', 'fr': "Quelle fete de decembre celebre la naissance de Jesus-Christ ?", 'da': 'Hvilken helligdag i december fejrer Jesu Kristi fødsel?'},
 'opt1': {'en': 'Easter', 'de': 'Ostern', 'es': 'Pascua', 'nl': 'Pasen', 'sv': 'Pask', 'fr': 'Pâques', 'da': 'Påske'},
 'opt2': {'en': 'Christmas', 'de': 'Weihnachten', 'es': 'Navidad', 'nl': 'Kerstmis', 'sv': 'Jul', 'fr': 'Noel', 'da': 'Jul'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which continent is home to the Amazon Rainforest's largest share?", 'de': 'Auf welchem Kontinent liegt der groesste Teil des Amazonas-Regenwaldes?', 'es': '¿En que continente esta la mayor parte de la selva amazonica?', 'nl': 'Op welk continent ligt het grootste deel van het Amazoneregenwoud?', 'sv': 'Pa vilken kontinent ligger den storsta delen av Amazonas regnskog?', 'fr': "Sur quel continent se trouve la plus grande part de la foret amazonienne ?", 'da': 'Hvilket kontinent huser den største andel af Amazonas regnskov?'},
 'opt1': {'en': 'South America', 'de': 'Suedamerika', 'es': 'Sudamerica', 'nl': 'Zuid-Amerika', 'sv': 'Sydamerika', 'fr': "l'Amerique du Sud", 'da': 'Sydamerika'},
 'opt2': {'en': 'Africa', 'de': 'Afrika', 'es': 'Africa', 'nl': 'Afrika', 'sv': 'Afrika', 'fr': "l'Afrique", 'da': 'Afrika'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which author wrote 'Alice's Adventures in Wonderland'?", 'de': "Welcher Autor schrieb 'Alice im Wunderland'?", 'es': "¿Que autor escribio 'Las aventuras de Alicia en el pais de las maravillas'?", 'nl': "Welke auteur schreef 'Alice's Avonturen in Wonderland'?", 'sv': "Vilken forfattare skrev 'Alice i Underlandet'?", 'fr': "Quel auteur a ecrit 'Les Aventures d'Alice au pays des merveilles' ?", 'da': "Hvilken forfatter skrev 'Alice i Eventyrland'?"},
 'opt1': {'en': 'J.M. Barrie', 'de': 'J.M. Barrie', 'es': 'J.M. Barrie', 'nl': 'J.M. Barrie', 'sv': 'J.M. Barrie', 'fr': 'J.M. Barrie', 'da': 'J.M. Barrie'},
 'opt2': {'en': 'Lewis Carroll', 'de': 'Lewis Carroll', 'es': 'Lewis Carroll', 'nl': 'Lewis Carroll', 'sv': 'Lewis Carroll', 'fr': 'Lewis Carroll', 'da': 'Lewis Carroll'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic toy is a doll with a boyfriend named Ken?', 'de': 'Welches klassische Spielzeug ist eine Puppe mit einem Freund namens Ken?', 'es': '¿Que juguete clasico es una muñeca con un novio llamado Ken?', 'nl': 'Welk klassiek speelgoed is een pop met een vriendje genaamd Ken?', 'sv': 'Vilken klassisk leksak ar en docka med en pojkvan vid namn Ken?', 'fr': "Quel jouet classique est une poupee avec un petit ami nomme Ken ?", 'da': 'Hvilket klassisk legetøj er en dukke med en kæreste ved navn Ken?'},
 'opt1': {'en': 'Barbie', 'de': 'Barbie', 'es': 'Barbie', 'nl': 'Barbie', 'sv': 'Barbie', 'fr': 'Barbie', 'da': 'Barbie'},
 'opt2': {'en': 'Bratz', 'de': 'Bratz', 'es': 'Bratz', 'nl': 'Bratz', 'sv': 'Bratz', 'fr': 'Bratz', 'da': 'Bratz'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which white dairy product is used to make cheese and yogurt?', 'de': 'Welches weisse Milchprodukt wird verwendet, um Kaese und Joghurt herzustellen?', 'es': '¿Que producto lacteo blanco se usa para hacer queso y yogur?', 'nl': 'Welk wit zuivelproduct wordt gebruikt om kaas en yoghurt te maken?', 'sv': 'Vilken vit mejeriprodukt anvands for att gora ost och yoghurt?', 'fr': "Quel produit laitier blanc est utilise pour fabriquer du fromage et du yaourt ?", 'da': 'Hvilket hvidt mejeriprodukt bruges til at lave ost og yoghurt?'},
 'opt1': {'en': 'Milk', 'de': 'Milch', 'es': 'Leche', 'nl': 'Melk', 'sv': 'Mjolk', 'fr': 'Lait', 'da': 'Mælk'},
 'opt2': {'en': 'Cream', 'de': 'Sahne', 'es': 'Crema', 'nl': 'Room', 'sv': 'Gradde', 'fr': 'Creme', 'da': 'Fløde'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. war ended slavery through the Emancipation Proclamation?', 'de': 'Welcher US-Krieg beendete die Sklaverei durch die Emanzipationsproklamation?', 'es': '¿Que guerra de EE.UU. termino con la esclavitud a traves de la Proclamacion de Emancipacion?', 'nl': 'Welke Amerikaanse oorlog beeindigde de slavernij door de Emancipatieproclamatie?', 'sv': 'Vilket amerikanskt krig avslutade slaveriet genom frigivningsproklamationen?', 'fr': "Quelle guerre americaine a mis fin a l'esclavage grace a la Proclamation d'emancipation ?", 'da': 'Hvilken amerikansk krig afsluttede slaveriet gennem frigørelsesproklamationen?'},
 'opt1': {'en': 'The Revolutionary War', 'de': 'Der Unabhaengigkeitskrieg', 'es': 'La Guerra Revolucionaria', 'nl': 'De Onafhankelijkheidsoorlog', 'sv': 'Revolutionskriget', 'fr': "La guerre d'Independance", 'da': 'Uafhængighedskrigen'},
 'opt2': {'en': 'The Civil War', 'de': 'Der Buergerkrieg', 'es': 'La Guerra Civil', 'nl': 'De Burgeroorlog', 'sv': 'Inbordeskriget', 'fr': 'La guerre de Secession', 'da': 'Borgerkrigen'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which decade is associated with the birth of hip hop music in the Bronx?', 'de': 'Mit welchem Jahrzehnt wird die Geburt der Hip-Hop-Musik in der Bronx assoziiert?', 'es': '¿Con que decada se asocia el nacimiento de la musica hip hop en el Bronx?', 'nl': 'Met welk decennium wordt de geboorte van hiphopmuziek in de Bronx geassocieerd?', 'sv': 'Vilket artionde forknippas med hiphopmusikens fodelse i Bronx?', 'fr': "A quelle decennie associe-t-on la naissance du hip-hop dans le Bronx ?", 'da': 'Hvilket årti forbindes med fødslen af hiphopmusik i Bronx?'},
 'opt1': {'en': '1970s', 'de': '1970er', 'es': 'los años 70', 'nl': 'jaren 70', 'sv': '1970-talet', 'fr': 'les annees 1970', 'da': '1970’erne'},
 'opt2': {'en': '1990s', 'de': '1990er', 'es': 'los años 90', 'nl': 'jaren 90', 'sv': '1990-talet', 'fr': 'les annees 1990', 'da': '1990’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which country is home to Niagara Falls' larger, horseshoe-shaped side?", 'de': 'In welchem Land liegt die groessere, hufeisenfoermige Seite der Niagarafaelle?', 'es': '¿En que pais esta el lado mas grande, con forma de herradura, de las Cataratas del Niagara?', 'nl': 'In welk land ligt de grotere, hoefijzervormige zijde van Niagara Falls?', 'sv': 'I vilket land ligger den storre, hastskoformade sidan av Niagarafallen?', 'fr': "Dans quel pays se trouve le cote le plus grand, en forme de fer a cheval, des chutes du Niagara ?", 'da': 'Hvilket land huser den større, hesteskoformede side af Niagara Falls?'},
 'opt1': {'en': 'United States', 'de': 'Vereinigte Staaten', 'es': 'Estados Unidos', 'nl': 'Verenigde Staten', 'sv': 'USA', 'fr': 'les États-Unis', 'da': 'USA'},
 'opt2': {'en': 'Canada', 'de': 'Kanada', 'es': 'Canada', 'nl': 'Canada', 'sv': 'Kanada', 'fr': 'le Canada', 'da': 'Canada'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which farm animal oinks and roots in mud?', 'de': 'Welches Nutztier grunzt und wuehlt im Schlamm?', 'es': '¿Que animal de granja gruñe y hurga en el barro?', 'nl': 'Welk boerderijdier knort en wroet in de modder?', 'sv': 'Vilket lantbruksdjur grymtar och botar i leran?', 'fr': "Quel animal de ferme grogne et fouille dans la boue ?", 'da': 'Hvilket landbrugsdyr grynter og rodfæster i mudder?'},
 'opt1': {'en': 'Goat', 'de': 'Ziege', 'es': 'Cabra', 'nl': 'Geit', 'sv': 'Get', 'fr': 'Chevre', 'da': 'Ged'},
 'opt2': {'en': 'Pig', 'de': 'Schwein', 'es': 'Cerdo', 'nl': 'Varken', 'sv': 'Gris', 'fr': 'Cochon', 'da': 'Gris'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which gemstone is the birthstone for those born in September?', 'de': 'Welcher Edelstein ist der Geburtsstein fuer im September Geborene?', 'es': '¿Que piedra preciosa es la piedra de nacimiento para los nacidos en septiembre?', 'nl': 'Welke edelsteen is de geboortesteen voor mensen geboren in september?', 'sv': 'Vilken adelsten ar fodelsestenen for de fodda i september?', 'fr': "Quelle pierre precieuse est la pierre de naissance de septembre ?", 'da': 'Hvilken ædelsten er fødselsstenen for dem, der er født i september?'},
 'opt1': {'en': 'Topaz', 'de': 'Topas', 'es': 'Topacio', 'nl': 'Topaas', 'sv': 'Topas', 'fr': 'Topaze', 'da': 'Topas'},
 'opt2': {'en': 'Sapphire', 'de': 'Saphir', 'es': 'Zafiro', 'nl': 'Saffier', 'sv': 'Safir', 'fr': 'Saphir', 'da': 'Safir'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic game show involves spinning a wheel to win cash and prizes?', 'de': 'Bei welcher klassischen Spielshow dreht man ein Rad, um Geld und Preise zu gewinnen?', 'es': '¿Que programa de concurso clasico implica girar una rueda para ganar dinero y premios?', 'nl': 'Bij welke klassieke spelshow draai je aan een rad om geld en prijzen te winnen?', 'sv': 'Vilken klassisk tavlingsshow innebar att snurra pa ett hjul for att vinna pengar och priser?', 'fr': "Quel jeu televise classique consiste a faire tourner une roue pour gagner de l'argent et des prix ?", 'da': 'Hvilket klassisk gameshow involverer at dreje et hjul for at vinde kontanter og præmier?'},
 'opt1': {'en': 'Wheel of Fortune', 'de': 'Gluecksrad', 'es': 'La Rueda de la Fortuna', 'nl': 'Rad van Fortuin', 'sv': 'Lyckohjulet', 'fr': 'La Roue de la fortune', 'da': 'Lykkehjulet'},
 'opt2': {'en': 'Jeopardy!', 'de': 'Jeopardy!', 'es': 'Jeopardy!', 'nl': 'Jeopardy!', 'sv': 'Jeopardy!', 'fr': 'Jeopardy!', 'da': 'Jeopardy!'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which red spice is made from dried and ground chili peppers, common in Hungarian cuisine?', 'de': 'Welches rote Gewuerz wird aus getrockneten und gemahlenen Chilischoten hergestellt und ist ueblich in der ungarischen Kueche?', 'es': '¿Que especia roja se hace con chiles secos y molidos, comun en la cocina hungara?', 'nl': 'Welke rode specerij wordt gemaakt van gedroogde en gemalen chilipepers, gebruikelijk in de Hongaarse keuken?', 'sv': 'Vilken rod krydda gors av torkade och malda chilipeppar, vanlig i ungersk mat?', 'fr': "Quelle epice rouge est faite de piments seches et moulus, courante dans la cuisine hongroise ?", 'da': 'Hvilket rødt krydderi er lavet af tørrede og malede chilipeber, almindeligt i ungarsk mad?'},
 'opt1': {'en': 'Cayenne', 'de': 'Cayennepfeffer', 'es': 'Cayena', 'nl': 'Cayennepeper', 'sv': 'Cayennepeppar', 'fr': 'Le cayenne', 'da': 'Cayennepeber'},
 'opt2': {'en': 'Paprika', 'de': 'Paprika', 'es': 'Pimenton', 'nl': 'Paprikapoeder', 'sv': 'Paprikapulver', 'fr': 'Le paprika', 'da': 'Paprika'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which flower is traditionally associated with Easter celebrations?', 'de': 'Welche Blume wird traditionell mit Osterfeiern assoziiert?', 'es': '¿Que flor se asocia tradicionalmente con las celebraciones de Pascua?', 'nl': 'Welke bloem wordt traditioneel geassocieerd met Paasvieringen?', 'sv': 'Vilken blomma forknippas traditionellt med paskfiranden?', 'fr': "Quelle fleur est traditionnellement associee aux celebrations de Pâques ?", 'da': 'Hvilken blomst forbindes traditionelt med påskefejringer?'},
 'opt1': {'en': 'Tulip', 'de': 'Tulpe', 'es': 'Tulipan', 'nl': 'Tulp', 'sv': 'Tulpan', 'fr': 'Tulipe', 'da': 'Tulipan'},
 'opt2': {'en': 'Lily', 'de': 'Lilie', 'es': 'Lirio', 'nl': 'Lelie', 'sv': 'Lilja', 'fr': 'Lys', 'da': 'Lilje'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Grand Ole Opry in Nashville?', 'de': 'In welchem US-Bundesstaat liegt die Grand Ole Opry in Nashville?', 'es': '¿En que estado de EE.UU. esta el Grand Ole Opry en Nashville?', 'nl': 'In welke Amerikaanse staat ligt de Grand Ole Opry in Nashville?', 'sv': 'I vilken amerikansk delstat ligger Grand Ole Opry i Nashville?', 'fr': "Dans quel État americain se trouve le Grand Ole Opry a Nashville ?", 'da': 'I hvilken amerikansk delstat ligger Grand Ole Opry i Nashville?'},
 'opt1': {'en': 'Tennessee', 'de': 'Tennessee', 'es': 'Tennessee', 'nl': 'Tennessee', 'sv': 'Tennessee', 'fr': 'le Tennessee', 'da': 'Tennessee'},
 'opt2': {'en': 'Kentucky', 'de': 'Kentucky', 'es': 'Kentucky', 'nl': 'Kentucky', 'sv': 'Kentucky', 'fr': 'le Kentucky', 'da': 'Kentucky'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves folding paper into decorative shapes, originating in Japan?', 'de': 'Bei welchem Handwerk wird Papier zu dekorativen Formen gefaltet, entstanden in Japan?', 'es': '¿Que manualidad implica doblar papel en formas decorativas, originaria de Japon?', 'nl': 'Bij welk handwerk vouw je papier in decoratieve vormen, afkomstig uit Japan?', 'sv': 'Vilket hantverk innebar att vika papper till dekorativa former, med ursprung i Japan?', 'fr': "Quel art manuel consiste a plier du papier en formes decoratives, originaire du Japon ?", 'da': 'Hvilket håndværk involverer at folde papir til dekorative former, med oprindelse i Japan?'},
 'opt1': {'en': 'Origami', 'de': 'Origami', 'es': 'Origami', 'nl': 'Origami', 'sv': 'Origami', 'fr': 'Origami', 'da': 'Origami'},
 'opt2': {'en': 'Kirigami', 'de': 'Kirigami', 'es': 'Kirigami', 'nl': 'Kirigami', 'sv': 'Kirigami', 'fr': 'Kirigami', 'da': 'Kirigami'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves two teams trying to get a puck into a net using sticks?', 'de': 'Bei welcher Sportart versuchen zwei Teams, einen Puck mit Schlaegern ins Netz zu bekommen?', 'es': '¿Que deporte implica que dos equipos intenten meter un disco en una porteria usando palos?', 'nl': 'Bij welke sport proberen twee teams een puck met stokken in een doel te krijgen?', 'sv': 'Vilken sport innebar att tva lag forsoker fa en puck i ett mal med klubbor?', 'fr': "Quel sport implique deux equipes essayant de mettre un palet dans un filet avec des batons ?", 'da': 'Hvilken sport involverer, at to hold prøver at få en puck i et mål med stave?'},
 'opt1': {'en': 'Ice hockey', 'de': 'Eishockey', 'es': 'Hockey sobre hielo', 'nl': 'IJshockey', 'sv': 'Ishockey', 'fr': 'Le hockey sur glace', 'da': 'Ishockey'},
 'opt2': {'en': 'Field hockey', 'de': 'Feldhockey', 'es': 'Hockey sobre cesped', 'nl': 'Veldhockey', 'sv': 'Landhockey', 'fr': 'Le hockey sur gazon', 'da': 'Feltløb'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which item of jewelry is worn around the finger, often to signify marriage?', 'de': 'Welcher Schmuck wird um den Finger getragen, oft als Zeichen der Ehe?', 'es': '¿Que joya se lleva alrededor del dedo, a menudo para significar matrimonio?', 'nl': 'Welk sieraad wordt om de vinger gedragen, vaak om huwelijk aan te duiden?', 'sv': 'Vilket smycke baers runt fingret, ofta for att beteckna aktenskap?', 'fr': "Quel bijou se porte autour du doigt, souvent pour signifier le mariage ?", 'da': 'Hvilket smykke bæres om fingeren, ofte for at betegne ægteskab?'},
 'opt1': {'en': 'Bracelet', 'de': 'Armband', 'es': 'Pulsera', 'nl': 'Armband', 'sv': 'Armband', 'fr': 'Bracelet', 'da': 'Armbånd'},
 'opt2': {'en': 'Ring', 'de': 'Ring', 'es': 'Anillo', 'nl': 'Ring', 'sv': 'Ring', 'fr': 'Bague', 'da': 'Ring'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic cartoon features a sailor with a spinach obsession?', 'de': 'In welchem klassischen Cartoon kommt ein Seemann mit einer Spinat-Obsession vor?', 'es': '¿Que dibujo animado clasico presenta a un marinero obsesionado con las espinacas?', 'nl': 'In welke klassieke tekenfilm komt een matroos voor die geobsedeerd is door spinazie?', 'sv': 'I vilken klassisk tecknad film forekommer en sjoman med en spenatobsession?', 'fr': "Quel dessin anime classique met en scene un marin obsede par les epinards ?", 'da': 'Hvilken klassisk tegnefilm har en sømand med en spinat-besættelse?'},
 'opt1': {'en': 'Popeye', 'de': 'Popeye', 'es': 'Popeye', 'nl': 'Popeye', 'sv': 'Karl-Alfred', 'fr': 'Popeye', 'da': 'Skipper Skræk'},
 'opt2': {'en': 'Bluto', 'de': 'Bluto', 'es': 'Bluto', 'nl': 'Bluto', 'sv': 'Bluto', 'fr': 'Brutus', 'da': 'Bluto'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the fjord city of Bergen?', 'de': 'In welchem Land liegt die Fjordstadt Bergen?', 'es': '¿En que pais esta la ciudad de fiordos Bergen?', 'nl': 'In welk land ligt de fjordstad Bergen?', 'sv': 'I vilket land ligger fjordstaden Bergen?', 'fr': "Dans quel pays se trouve la ville de fjord Bergen ?", 'da': 'Hvilket land huser fjordbyen Bergen?'},
 'opt1': {'en': 'Sweden', 'de': 'Schweden', 'es': 'Suecia', 'nl': 'Zweden', 'sv': 'Sverige', 'fr': 'la Suede', 'da': 'Sverige'},
 'opt2': {'en': 'Norway', 'de': 'Norwegen', 'es': 'Noruega', 'nl': 'Noorwegen', 'sv': 'Norge', 'fr': 'la Norvège', 'da': 'Norge'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady worked as a photojournalist before marrying into politics?', 'de': 'Welche First Lady arbeitete als Fotojournalistin, bevor sie in die Politik einheiratete?', 'es': '¿Que primera dama trabajo como fotoperiodista antes de casarse con un politico?', 'nl': 'Welke First Lady werkte als fotojournalist voordat ze in de politiek trouwde?', 'sv': 'Vilken First Lady arbetade som fotojournalist innan hon gifte in i politiken?', 'fr': "Quelle Première dame a travaille comme photojournaliste avant d'epouser un homme politique ?", 'da': 'Hvilken førstedame arbejdede som fotojournalist, før hun giftede sig ind i politik?'},
 'opt1': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'},
 'opt2': {'en': 'Pat Nixon', 'de': 'Pat Nixon', 'es': 'Pat Nixon', 'nl': 'Pat Nixon', 'sv': 'Pat Nixon', 'fr': 'Pat Nixon', 'da': 'Pat Nixon'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Peruvian dish consists of raw fish marinated in citrus juice?', 'de': 'Welches peruanische Gericht besteht aus rohem Fisch, mariniert in Zitrussaft?', 'es': '¿Que plato peruano consiste en pescado crudo marinado en jugo citrico?', 'nl': 'Welk Peruaans gerecht bestaat uit rauwe vis gemarineerd in citrussap?', 'sv': 'Vilken peruansk ratt bestar av ra fisk marinerad i citrusjuice?', 'fr': "Quel plat peruvien consiste en du poisson cru marine dans du jus d'agrumes ?", 'da': 'Hvilken peruviansk ret består af rå fisk marineret i citrussaft?'},
 'opt1': {'en': 'Ceviche', 'de': 'Ceviche', 'es': 'Ceviche', 'nl': 'Ceviche', 'sv': 'Ceviche', 'fr': 'Ceviche', 'da': 'Ceviche'},
 'opt2': {'en': 'Tiradito', 'de': 'Tiradito', 'es': 'Tiradito', 'nl': 'Tiradito', 'sv': 'Tiradito', 'fr': 'Tiradito', 'da': 'Tiradito'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which gas makes up the majority of Earth's atmosphere?", 'de': 'Welches Gas macht den Großteil der Erdatmosphaere aus?', 'es': '¿Que gas constituye la mayor parte de la atmosfera terrestre?', 'nl': 'Welk gas vormt het grootste deel van de aardatmosfeer?', 'sv': 'Vilken gas utgor storsta delen av jordens atmosfar?', 'fr': "Quel gaz constitue la majeure partie de l'atmosphere terrestre ?", 'da': 'Hvilken gas udgør størstedelen af Jordens atmosfære?'},
 'opt1': {'en': 'Nitrogen', 'de': 'Stickstoff', 'es': 'Nitrogeno', 'nl': 'Stikstof', 'sv': 'Kvave', 'fr': 'Azote', 'da': 'Kvælstof'},
 'opt2': {'en': 'Oxygen', 'de': 'Sauerstoff', 'es': 'Oxigeno', 'nl': 'Zuurstof', 'sv': 'Syre', 'fr': 'Oxygene', 'da': 'Ilt'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'Frankenstein'?", 'de': "Welche Autorin schrieb 'Frankenstein'?", 'es': "¿Que autora escribio 'Frankenstein'?", 'nl': "Welke auteur schreef 'Frankenstein'?", 'sv': "Vilken forfattare skrev 'Frankenstein'?", 'fr': "Quelle autrice a ecrit 'Frankenstein' ?", 'da': "Hvilken forfatter skrev 'Frankenstein'?"},
 'opt1': {'en': 'Bram Stoker', 'de': 'Bram Stoker', 'es': 'Bram Stoker', 'nl': 'Bram Stoker', 'sv': 'Bram Stoker', 'fr': 'Bram Stoker', 'da': 'Bram Stoker'},
 'opt2': {'en': 'Mary Shelley', 'de': 'Mary Shelley', 'es': 'Mary Shelley', 'nl': 'Mary Shelley', 'sv': 'Mary Shelley', 'fr': 'Mary Shelley', 'da': 'Mary Shelley'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Elphaba in the original Broadway cast of 'Wicked'?", 'de': "Welche Schauspielerin spielte Elphaba in der urspruenglichen Broadway-Besetzung von 'Wicked'?", 'es': "¿Que actriz interpreto a Elphaba en el elenco original de Broadway de 'Wicked'?", 'nl': "Welke actrice speelde Elphaba in de originele Broadway-cast van 'Wicked'?", 'sv': "Vilken skadespelerska spelade Elphaba i den ursprungliga Broadway-uppsattningen av 'Wicked'?", 'fr': "Quelle actrice a joue Elphaba dans la distribution originale de Broadway de 'Wicked' ?", 'da': "Hvilken skuespiller spillede Elphaba i den originale Broadway-opsætning af 'Wicked'?"},
 'opt1': {'en': 'Kristin Chenoweth', 'de': 'Kristin Chenoweth', 'es': 'Kristin Chenoweth', 'nl': 'Kristin Chenoweth', 'sv': 'Kristin Chenoweth', 'fr': 'Kristin Chenoweth', 'da': 'Kristin Chenoweth'},
 'opt2': {'en': 'Idina Menzel', 'de': 'Idina Menzel', 'es': 'Idina Menzel', 'nl': 'Idina Menzel', 'sv': 'Idina Menzel', 'fr': 'Idina Menzel', 'da': 'Idina Menzel'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which singer is known for the song 'Respect' and civil rights advocacy?", 'de': "Welche Saengerin ist fuer das Lied 'Respect' und ihr Eintreten fuer Buergerrechte bekannt?", 'es': "¿Que cantante es conocida por la cancion 'Respect' y su defensa de los derechos civiles?", 'nl': "Welke zangeres staat bekend om het lied 'Respect' en pleiten voor burgerrechten?", 'sv': "Vilken sangerska ar kand for latten 'Respect' och for medborgarrattskampen?", 'fr': "Quelle chanteuse est connue pour la chanson 'Respect' et son engagement pour les droits civiques ?", 'da': "Hvilken sangerinde er kendt for sangen 'Respect' og fortalervirksomhed for borgerrettigheder?"},
 'opt1': {'en': 'Aretha Franklin', 'de': 'Aretha Franklin', 'es': 'Aretha Franklin', 'nl': 'Aretha Franklin', 'sv': 'Aretha Franklin', 'fr': 'Aretha Franklin', 'da': 'Aretha Franklin'},
 'opt2': {'en': 'Nina Simone', 'de': 'Nina Simone', 'es': 'Nina Simone', 'nl': 'Nina Simone', 'sv': 'Nina Simone', 'fr': 'Nina Simone', 'da': 'Nina Simone'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient ruins of Palenque?', 'de': 'In welchem Land liegen die antiken Ruinen von Palenque?', 'es': '¿En que pais estan las ruinas antiguas de Palenque?', 'nl': 'In welk land liggen de oude ruines van Palenque?', 'sv': 'I vilket land ligger de antika ruinerna av Palenque?', 'fr': "Dans quel pays se trouvent les ruines antiques de Palenque ?", 'da': 'Hvilket land huser de gamle ruiner af Palenque?'},
 'opt1': {'en': 'Mexico', 'de': 'Mexiko', 'es': 'Mexico', 'nl': 'Mexico', 'sv': 'Mexiko', 'fr': 'le Mexique', 'da': 'Mexico'},
 'opt2': {'en': 'Belize', 'de': 'Belize', 'es': 'Belice', 'nl': 'Belize', 'sv': 'Belize', 'fr': 'le Belize', 'da': 'Belize'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which large Australian marsupial carries its young in a pouch and hops?', 'de': 'Welches grosse australische Beuteltier traegt seine Jungen in einem Beutel und huepft?', 'es': '¿Que gran marsupial australiano lleva a sus crias en una bolsa y salta?', 'nl': 'Welk groot Australisch buideldier draagt zijn jongen in een buidel en huppelt?', 'sv': 'Vilket stort australiskt pungdjur baer sina ungar i en pung och hoppar?', 'fr': "Quel grand marsupial australien porte ses petits dans une poche et bondit ?", 'da': 'Hvilket stort australsk pungdyr bærer sine unger i en pung og hopper?'},
 'opt1': {'en': 'Wallaby', 'de': 'Wallaby', 'es': 'Walabi', 'nl': 'Wallaby', 'sv': 'Wallaby', 'fr': 'Wallaby', 'da': 'Wallaby'},
 'opt2': {'en': 'Kangaroo', 'de': 'Kaenguru', 'es': 'Canguro', 'nl': 'Kangoeroe', 'sv': 'Kangaru', 'fr': 'Kangourou', 'da': 'Kænguru'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Moroccan dish is a slow-cooked stew named after its conical clay pot?', 'de': 'Welches marokkanische Gericht ist ein langsam gekochter Eintopf, benannt nach seinem kegelfoermigen Tontopf?', 'es': '¿Que plato marroqui es un guiso de coccion lenta llamado por su olla conica de barro?', 'nl': 'Welk Marokkaans gerecht is een langzaam gegaarde stoofpot vernoemd naar zijn kegelvormige kleipot?', 'sv': 'Vilken marockansk ratt ar en langsamt tillagad gryta uppkallad efter sin koniska lerkruka?', 'fr': "Quel plat marocain est un ragout mijote nomme d'apres son plat conique en argile ?", 'da': 'Hvilken marokkansk ret er en langsomt tilberedt gryde opkaldt efter sin kegleformede lerkrukke?'},
 'opt1': {'en': 'Couscous', 'de': 'Couscous', 'es': 'Cuscus', 'nl': 'Couscous', 'sv': 'Couscous', 'fr': 'Couscous', 'da': 'Couscous'},
 'opt2': {'en': 'Tagine', 'de': 'Tajine', 'es': 'Tajine', 'nl': 'Tajine', 'sv': 'Tagine', 'fr': 'Tajine', 'da': 'Tagine'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which First Lady founded a organization to preserve historic White House furnishings?', 'de': 'Welche First Lady gruendete eine Organisation zum Erhalt historischer Einrichtungsgegenstaende des Weissen Hauses?', 'es': '¿Que primera dama fundo una organizacion para preservar el mobiliario historico de la Casa Blanca?', 'nl': 'Welke First Lady richtte een organisatie op om historisch meubilair van het Witte Huis te behouden?', 'sv': 'Vilken First Lady grundade en organisation for att bevara Vita husets historiska inredning?', 'fr': "Quelle Première dame a fonde une organisation pour preserver le mobilier historique de la Maison-Blanche ?", 'da': 'Hvilken førstedame grundlagde en organisation til at bevare Det Hvide Hus’ historiske inventar?'},
 'opt1': {'en': 'Edith Roosevelt', 'de': 'Edith Roosevelt', 'es': 'Edith Roosevelt', 'nl': 'Edith Roosevelt', 'sv': 'Edith Roosevelt', 'fr': 'Edith Roosevelt', 'da': 'Edith Roosevelt'},
 'opt2': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which artist is known for large-scale Campbell's soup can silkscreens?", 'de': "Welcher Kuenstler ist fuer grossformatige Siebdrucke von Campbell's-Suppendosen bekannt?", 'es': "¿Que artista es conocido por serigrafias a gran escala de latas de sopa Campbell's?", 'nl': "Welke kunstenaar staat bekend om grootschalige zeefdrukken van Campbell's-soepblikken?", 'sv': "Vilken konstnar ar kand for storskaliga silkscreentryck av Campbell's soppburkar?", 'fr': "Quel artiste est connu pour ses serigraphies grand format de boites de soupe Campbell's ?", 'da': "Hvilken kunstner er kendt for store silketryk af Campbell's suppedåser?"},
 'opt1': {'en': 'Jasper Johns', 'de': 'Jasper Johns', 'es': 'Jasper Johns', 'nl': 'Jasper Johns', 'sv': 'Jasper Johns', 'fr': 'Jasper Johns', 'da': 'Jasper Johns'},
 'opt2': {'en': 'Andy Warhol', 'de': 'Andy Warhol', 'es': 'Andy Warhol', 'nl': 'Andy Warhol', 'sv': 'Andy Warhol', 'fr': 'Andy Warhol', 'da': 'Andy Warhol'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which flowering plant is known for its large, colorful blooms and is Japan's cherry blossom?", 'de': 'Welche bluehende Pflanze ist fuer ihre grossen, farbenfrohen Blueten bekannt und ist Japans Kirschbluete?', 'es': '¿Que planta con flores es conocida por sus grandes y coloridas flores y es el cerezo japones?', 'nl': 'Welke bloeiende plant staat bekend om zijn grote, kleurrijke bloesems en is de kersenbloesem van Japan?', 'sv': 'Vilken blommande vaxt ar kand for sina stora, fargglada blommor och ar Japans korsbarsblom?', 'fr': "Quelle plante a fleurs est connue pour ses grandes fleurs colorees et est la fleur de cerisier du Japon ?", 'da': 'Hvilken blomstrende plante er kendt for sine store, farverige blomster og er Japans kirsebærblomst?'},
 'opt1': {'en': 'Sakura', 'de': 'Sakura', 'es': 'Sakura', 'nl': 'Sakura', 'sv': 'Sakura', 'fr': 'Sakura', 'da': 'Sakura'},
 'opt2': {'en': 'Plum blossom', 'de': 'Pflaumenbluete', 'es': 'Flor de ciruelo', 'nl': 'Pruimenbloesem', 'sv': 'Plommonblom', 'fr': 'Fleur de prunier', 'da': 'Blommeblomst'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which musical is based on Victor Hugo's novel about a Parisian revolution?", 'de': "Welches Musical basiert auf Victor Hugos Roman ueber eine Pariser Revolution?", 'es': "¿Que musical se basa en la novela de Victor Hugo sobre una revolucion parisina?", 'nl': "Welke musical is gebaseerd op Victor Hugo's roman over een Parijse revolutie?", 'sv': "Vilken musikal ar baserad pa Victor Hugos roman om en parisisk revolution?", 'fr': "Quelle comedie musicale s'inspire du roman de Victor Hugo sur une revolution parisienne ?", 'da': 'Hvilken musical er baseret på Victor Hugos roman om en parisisk revolution?'},
 'opt1': {'en': 'The Phantom of the Opera', 'de': 'Das Phantom der Oper', 'es': 'El Fantasma de la Opera', 'nl': 'Het Spook van de Opera', 'sv': 'Fantomen pa Operan', 'fr': "Le Fantome de l'Opera", 'da': 'Fantomet i Operaen'},
 'opt2': {'en': 'Les Miserables', 'de': 'Les Miserables', 'es': 'Los Miserables', 'nl': 'Les Miserables', 'sv': 'Les Miserables', 'fr': 'Les Miserables', 'da': 'Les Miserables'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country has Budapest as its capital, split by the Danube River?', 'de': 'Welches Land hat Budapest als Hauptstadt, geteilt durch die Donau?', 'es': '¿Que pais tiene a Budapest como su capital, dividida por el rio Danubio?', 'nl': 'Welk land heeft Boedapest als hoofdstad, gesplitst door de rivier de Donau?', 'sv': 'Vilket land har Budapest som huvudstad, delad av floden Donau?', 'fr': "Quel pays a Budapest pour capitale, divisee par le Danube ?", 'da': 'Hvilket land har Budapest som hovedstad, delt af Donau-floden?'},
 'opt1': {'en': 'Hungary', 'de': 'Ungarn', 'es': 'Hungria', 'nl': 'Hongarije', 'sv': 'Ungern', 'fr': 'la Hongrie', 'da': 'Ungarn'},
 'opt2': {'en': 'Romania', 'de': 'Rumaenien', 'es': 'Rumania', 'nl': 'Roemenie', 'sv': 'Rumanien', 'fr': 'la Roumanie', 'da': 'Rumænien'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which spice is derived from tree bark and used in baking, especially with apples?', 'de': 'Welches Gewuerz stammt von Baumrinde und wird beim Backen verwendet, besonders mit Aepfeln?', 'es': '¿Que especia se deriva de la corteza de un arbol y se usa en reposteria, especialmente con manzanas?', 'nl': 'Welke specerij komt van boombast en wordt gebruikt bij het bakken, vooral met appels?', 'sv': 'Vilken krydda kommer fran trabark och anvands vid bakning, sarskilt med applen?', 'fr': "Quelle epice provient de l'ecorce d'arbre et est utilisee en patisserie, notamment avec les pommes ?", 'da': 'Hvilket krydderi kommer fra træbark og bruges i bagning, især med æbler?'},
 'opt1': {'en': 'Nutmeg', 'de': 'Muskatnuss', 'es': 'Nuez moscada', 'nl': 'Nootmuskaat', 'sv': 'Muskot', 'fr': 'La muscade', 'da': 'Muskatnød'},
 'opt2': {'en': 'Cinnamon', 'de': 'Zimt', 'es': 'Canela', 'nl': 'Kaneel', 'sv': 'Kanel', 'fr': 'La cannelle', 'da': 'Kanel'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which needlework craft uses a hoop to stretch fabric while stitching decorative designs?', 'de': 'Welches Handarbeitshandwerk verwendet einen Rahmen, um Stoff zu spannen, waehrend dekorative Muster gestickt werden?', 'es': '¿Que arte de aguja usa un aro para estirar la tela mientras se cosen diseños decorativos?', 'nl': 'Welk naaiwerk gebruikt een hoepel om stof te spannen tijdens het stikken van decoratieve patronen?', 'sv': 'Vilket handarbete anvander en bage for att stracka tyg medan man syr dekorativa monster?', 'fr': "Quel art de l'aiguille utilise un tambour pour tendre le tissu tout en brodant des motifs decoratifs ?", 'da': 'Hvilket håndarbejde bruger en ring til at strække stof, mens man syr dekorative mønstre?'},
 'opt1': {'en': 'Embroidery', 'de': 'Stickerei', 'es': 'Bordado', 'nl': 'Borduurwerk', 'sv': 'Broderi', 'fr': 'Broderie', 'da': 'Broderi'},
 'opt2': {'en': 'Quilting', 'de': 'Quilten', 'es': 'Acolchado', 'nl': 'Quilten', 'sv': 'Lapptacksarbete', 'fr': 'Le patchwork', 'da': 'Quiltning'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which organ pumps blood and is located in the chest?', 'de': 'Welches Organ pumpt Blut und liegt in der Brust?', 'es': '¿Que organo bombea sangre y esta ubicado en el pecho?', 'nl': 'Welk orgaan pompt bloed en bevindt zich in de borst?', 'sv': 'Vilket organ pumpar blod och sitter i brostet?', 'fr': "Quel organe pompe le sang et se situe dans la poitrine ?", 'da': 'Hvilket organ pumper blod og er placeret i brystet?'},
 'opt1': {'en': 'Lungs', 'de': 'Lunge', 'es': 'Pulmones', 'nl': 'Longen', 'sv': 'Lungor', 'fr': 'Poumons', 'da': 'Lunger'},
 'opt2': {'en': 'Heart', 'de': 'Herz', 'es': 'Corazon', 'nl': 'Hart', 'sv': 'Hjarta', 'fr': 'Coeur', 'da': 'Hjerte'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'In gymnastics, which apparatus involves swinging and releasing grips mid-air?', 'de': 'Bei welchem Turngeraet schwingt und loest man Griffe in der Luft?', 'es': '¿En gimnasia, que aparato implica balancearse y soltar agarres en el aire?', 'nl': 'Bij welk gymnastiektoestel zwaai je en laat je grepen los in de lucht?', 'sv': 'Vilken gymnastikapparat innebar att svanga och slappa tag i luften?', 'fr': "En gymnastique, quel agres implique de se balancer et de lacher les prises en l'air ?", 'da': 'I gymnastik, hvilket redskab involverer at svinge og slippe grebet midt i luften?'},
 'opt1': {'en': 'Uneven bars', 'de': 'Stufenbarren', 'es': 'Barras asimetricas', 'nl': 'Ongelijke brug', 'sv': 'Ojamna barr', 'fr': 'Barres asymetriques', 'da': 'Ujævne barrer'},
 'opt2': {'en': 'Balance beam', 'de': 'Schwebebalken', 'es': 'Barra de equilibrio', 'nl': 'Balk', 'sv': 'Bom', 'fr': 'Poutre', 'da': 'Bomme'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which German city is famous for its concentration of Baroque and Rococo architecture, nicknamed 'Florence on the Elbe'?", 'de': "Welche deutsche Stadt ist beruehmt fuer ihre Konzentration an Barock- und Rokoko-Architektur, genannt 'Elbflorenz'?", 'es': "¿Que ciudad alemana es famosa por su concentracion de arquitectura barroca y rococo, apodada 'Florencia del Elba'?", 'nl': "Welke Duitse stad staat bekend om zijn concentratie van barok- en rococo-architectuur, bijgenaamd 'Florence aan de Elbe'?", 'sv': "Vilken tysk stad ar kand for sin koncentration av barock- och rokokoarkitektur, med smeknamnet 'Elbflorens'?", 'fr': "Quelle ville allemande est celebre pour sa concentration d'architecture baroque et rococo, surnommee la 'Florence de l'Elbe' ?", 'da': "Hvilken tysk by er berømt for sin koncentration af barok- og rokokoarkitektur, kaldet 'Elben-Firenze'?"},
 'opt1': {'en': 'Dresden', 'de': 'Dresden', 'es': 'Dresde', 'nl': 'Dresden', 'sv': 'Dresden', 'fr': 'Dresde', 'da': 'Dresden'},
 'opt2': {'en': 'Berlin', 'de': 'Berlin', 'es': 'Berlin', 'nl': 'Berlijn', 'sv': 'Berlin', 'fr': 'Berlin', 'da': 'Berlin'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Belle in Disney's live-action 'Beauty and the Beast'?", 'de': "Welche Schauspielerin spielte Belle in Disneys Realfilm 'Die Schoene und das Biest'?", 'es': "¿Que actriz interpreto a Bella en la pelicula de accion real de Disney 'La Bella y la Bestia'?", 'nl': "Welke actrice speelde Belle in Disney's live-action 'Belle en het Beest'?", 'sv': "Vilken skadespelerska spelade Belle i Disneys spelfilmsversion av 'Skonheten och odjuret'?", 'fr': "Quelle actrice a joue Belle dans le film en prises de vues reelles 'La Belle et la Bete' de Disney ?", 'da': "Hvilken skuespiller spillede Belle i Disneys live-action 'Skønheden og Udyret'?"},
 'opt1': {'en': 'Emma Watson', 'de': 'Emma Watson', 'es': 'Emma Watson', 'nl': 'Emma Watson', 'sv': 'Emma Watson', 'fr': 'Emma Watson', 'da': 'Emma Watson'},
 'opt2': {'en': 'Emma Stone', 'de': 'Emma Stone', 'es': 'Emma Stone', 'nl': 'Emma Stone', 'sv': 'Emma Stone', 'fr': 'Emma Stone', 'da': 'Emma Stone'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which classic novel features the character Heathcliff?', 'de': 'In welchem Klassiker kommt die Figur Heathcliff vor?', 'es': '¿Que novela clasica presenta al personaje Heathcliff?', 'nl': 'In welke klassieke roman komt het personage Heathcliff voor?', 'sv': 'I vilken klassisk roman forekommer karaktaren Heathcliff?', 'fr': "Quel roman classique met en scene le personnage de Heathcliff ?", 'da': 'Hvilken klassisk roman har karakteren Heathcliff?'},
 'opt1': {'en': 'Wuthering Heights', 'de': 'Sturmhoehe', 'es': 'Cumbres Borrascosas', 'nl': 'Woeste Hoogten', 'sv': 'Svindlande hojder', 'fr': 'Les Hauts de Hurlevent', 'da': 'Stormfulde højder'},
 'opt2': {'en': 'Jane Eyre', 'de': 'Jane Eyre', 'es': 'Jane Eyre', 'nl': 'Jane Eyre', 'sv': 'Jane Eyre', 'fr': 'Jane Eyre', 'da': 'Jane Eyre'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which 1990s girl group sang 'No Scrubs'?", 'de': "Welche Girlgroup der 1990er sang 'No Scrubs'?", 'es': "¿Que grupo femenino de los 90 canto 'No Scrubs'?", 'nl': "Welke meidengroep uit de jaren 90 zong 'No Scrubs'?", 'sv': "Vilken tjejgrupp fran 1990-talet sjong 'No Scrubs'?", 'fr': "Quel girls band des annees 1990 a chante 'No Scrubs' ?", 'da': "Hvilken 1990’er-pigegruppe sang 'No Scrubs'?"},
 'opt1': {'en': 'TLC', 'de': 'TLC', 'es': 'TLC', 'nl': 'TLC', 'sv': 'TLC', 'fr': 'TLC', 'da': 'TLC'},
 'opt2': {'en': 'En Vogue', 'de': 'En Vogue', 'es': 'En Vogue', 'nl': 'En Vogue', 'sv': 'En Vogue', 'fr': 'En Vogue', 'da': 'En Vogue'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is made from cornmeal mush, fried or baked?', 'de': 'Welches suddstaatliche Gericht wird aus Maisbrei hergestellt, gebraten oder gebacken?', 'es': '¿Que plato sureño de EE.UU. se hace con papilla de maiz, frita o horneada?', 'nl': 'Welk zuidelijk Amerikaans gerecht wordt gemaakt van maispap, gefrituurd of gebakken?', 'sv': 'Vilken sydstatlig ratt gors av majsgrot, friterad eller bakad?', 'fr': "Quel plat du Sud des États-Unis est fait de bouillie de mais, frite ou cuite au four ?", 'da': 'Hvilken sydstatsret er lavet af majsgrød, stegt eller bagt?'},
 'opt1': {'en': 'Grits', 'de': 'Grits', 'es': 'Grits', 'nl': 'Grits', 'sv': 'Grits', 'fr': 'Grits', 'da': 'Grits'},
 'opt2': {'en': 'Polenta', 'de': 'Polenta', 'es': 'Polenta', 'nl': 'Polenta', 'sv': 'Polenta', 'fr': 'Polenta', 'da': 'Polenta'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which queen is known for uniting Spain with her husband Ferdinand and funding Columbus's voyage?", 'de': 'Welche Koenigin ist bekannt fuer die Vereinigung Spaniens mit ihrem Ehemann Ferdinand und die Finanzierung von Kolumbus\' Reise?', 'es': '¿Que reina es conocida por unir España con su esposo Fernando y financiar el viaje de Colon?', 'nl': 'Welke koningin staat bekend om het verenigen van Spanje met haar echtgenoot Ferdinand en het financieren van Columbus\' reis?', 'sv': 'Vilken drottning ar kand for att ha forenat Spanien med sin make Ferdinand och finansierat Columbus resa?', 'fr': "Quelle reine est connue pour avoir uni l'Espagne avec son mari Ferdinand et finance le voyage de Colomb ?", 'da': 'Hvilken dronning er kendt for at forene Spanien med sin mand Ferdinand og finansiere Columbus’ rejse?'},
 'opt1': {'en': 'Catherine of Aragon', 'de': 'Katharina von Aragon', 'es': 'Catalina de Aragon', 'nl': 'Catharina van Aragon', 'sv': 'Katarina av Aragonien', 'fr': "Catherine d'Aragon", 'da': 'Katarina af Aragonien'},
 'opt2': {'en': 'Isabella I of Castile', 'de': 'Isabella I. von Kastilien', 'es': 'Isabel I de Castilla', 'nl': 'Isabella I van Castilie', 'sv': 'Isabella I av Kastilien', 'fr': 'Isabelle Ire de Castille', 'da': 'Isabella 1. af Castilien'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large-eared African mammal has the largest ears relative to its body among canines?', 'de': 'Welches grossohrige afrikanische Saeugetier hat die groessten Ohren im Verhaeltnis zu seinem Koerper unter den Caniden?', 'es': '¿Que mamifero africano de orejas grandes tiene las orejas mas grandes en relacion con su cuerpo entre los caninos?', 'nl': 'Welk Afrikaans zoogdier met grote oren heeft de grootste oren in verhouding tot zijn lichaam onder de hondachtigen?', 'sv': 'Vilket afrikanskt daggdjur med stora oron har storst oron i forhallande till kroppen bland hunddjur?', 'fr': "Quel mammifere africain aux grandes oreilles a les oreilles les plus grandes par rapport a son corps parmi les canides ?", 'da': 'Hvilket afrikansk pattedyr med store ører har de største ører i forhold til kroppen blandt hundedyr?'},
 'opt1': {'en': 'Fennec fox', 'de': 'Fennek', 'es': 'Zorro fenec', 'nl': 'Fennek', 'sv': 'Ökenrav', 'fr': 'Le fennec', 'da': 'Fennekræv'},
 'opt2': {'en': 'African wild dog', 'de': 'Afrikanischer Wildhund', 'es': 'Perro salvaje africano', 'nl': 'Afrikaanse wilde hond', 'sv': 'Afrikansk vildhund', 'fr': 'Le lycaon', 'da': 'Afrikansk vildhund'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Persepolis?', 'de': 'In welchem Land liegt die antike Stadt Persepolis?', 'es': '¿En que pais esta la antigua ciudad de Persepolis?', 'nl': 'In welk land ligt de oude stad Persepolis?', 'sv': 'I vilket land ligger den antika staden Persepolis?', 'fr': "Dans quel pays se trouve l'ancienne cite de Persepolis ?", 'da': 'Hvilket land huser den antikke by Persepolis?'},
 'opt1': {'en': 'Iraq', 'de': 'Irak', 'es': 'Irak', 'nl': 'Irak', 'sv': 'Irak', 'fr': "l'Irak", 'da': 'Irak'},
 'opt2': {'en': 'Iran', 'de': 'Iran', 'es': 'Iran', 'nl': 'Iran', 'sv': 'Iran', 'fr': "l'Iran", 'da': 'Iran'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which type of shoe, popular in the 1970s, has a thick raised sole?', 'de': 'Welche Art von Schuh, populaer in den 1970er Jahren, hat eine dicke erhoehte Sohle?', 'es': '¿Que tipo de zapato, popular en los años 70, tiene una suela gruesa y elevada?', 'nl': 'Welk type schoen, populair in de jaren 70, heeft een dikke, verhoogde zool?', 'sv': 'Vilken typ av sko, popular pa 1970-talet, har en tjock upphojd sula?', 'fr': "Quel type de chaussure, populaire dans les annees 1970, a une semelle epaisse et surelevee ?", 'da': 'Hvilken type sko, populær i 1970’erne, har en tyk hævet sål?'},
 'opt1': {'en': 'Wedge', 'de': 'Keilabsatz', 'es': 'Cuña', 'nl': 'Sleehak', 'sv': 'Kilklack', 'fr': 'Compensee', 'da': 'Kilehæl'},
 'opt2': {'en': 'Platform shoe', 'de': 'Plateauschuh', 'es': 'Zapato de plataforma', 'nl': 'Plateauschoen', 'sv': 'Platformssko', 'fr': 'Chaussure a plateforme', 'da': 'Platformssko'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Netflix series follows a group of teens investigating supernatural events in Hawkins, Indiana?', 'de': 'Welche Netflix-Serie folgt einer Gruppe von Teenagern, die uebernatuerliche Ereignisse in Hawkins, Indiana, untersuchen?', 'es': '¿Que serie de Netflix sigue a un grupo de adolescentes investigando eventos sobrenaturales en Hawkins, Indiana?', 'nl': 'Welke Netflix-serie volgt een groep tieners die bovennatuurlijke gebeurtenissen onderzoeken in Hawkins, Indiana?', 'sv': 'Vilken Netflix-serie foljer en grupp tonaringar som utreder overnaturliga handelser i Hawkins, Indiana?', 'fr': "Quelle serie Netflix suit un groupe d'adolescents enquetant sur des evenements surnaturels a Hawkins, Indiana ?", 'da': 'Hvilken Netflix-serie følger en gruppe teenagere, der undersøger overnaturlige begivenheder i Hawkins, Indiana?'},
 'opt1': {'en': 'Dark', 'de': 'Dark', 'es': 'Dark', 'nl': 'Dark', 'sv': 'Dark', 'fr': 'Dark', 'da': 'Dark'},
 'opt2': {'en': 'Stranger Things', 'de': 'Stranger Things', 'es': 'Stranger Things', 'nl': 'Stranger Things', 'sv': 'Stranger Things', 'fr': 'Stranger Things', 'da': 'Stranger Things'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which tree is tapped for its sap to make rubber?', 'de': 'Welcher Baum wird angezapft, um Kautschuk aus seinem Saft zu gewinnen?', 'es': '¿Que arbol se sangra por su savia para hacer caucho?', 'nl': 'Welke boom wordt aangetapt voor zijn sap om rubber te maken?', 'sv': 'Vilket trad tappas pa sin sav for att gora gummi?', 'fr': "Quel arbre est saigne pour sa seve afin de fabriquer du caoutchouc ?", 'da': 'Hvilket træ tappes for sin saft til at lave gummi?'},
 'opt1': {'en': 'Sugar maple', 'de': 'Zuckerahorn', 'es': 'Arce azucarero', 'nl': 'Suikeresdoorn', 'sv': 'Sockerlonn', 'fr': 'Erable a sucre', 'da': 'Sukkerær'},
 'opt2': {'en': 'Rubber tree', 'de': 'Kautschukbaum', 'es': 'Arbol del caucho', 'nl': 'Rubberboom', 'sv': 'Gummitrad', 'fr': 'Hevea', 'da': 'Gummitræ'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Thai dish is a stir-fried noodle dish with shrimp, egg, and peanuts?', 'de': 'Welches thailaendische Gericht ist ein gebratenes Nudelgericht mit Garnelen, Ei und Erdnuessen?', 'es': '¿Que plato tailandes es un salteado de fideos con camarones, huevo y cacahuates?', 'nl': 'Welk Thais gerecht is een gewokte noedelgerecht met garnalen, ei en pinda\'s?', 'sv': 'Vilken thailandsk ratt ar en wokad nudelratt med raga, agg och jordnotter?', 'fr': "Quel plat thai est un plat de nouilles sautees aux crevettes, œuf et cacahuetes ?", 'da': 'Hvilken thailandsk ret er en wokket nudelret med rejer, æg og jordnødder?'},
 'opt1': {'en': 'Pad See Ew', 'de': 'Pad See Ew', 'es': 'Pad See Ew', 'nl': 'Pad See Ew', 'sv': 'Pad See Ew', 'fr': 'Pad see ew', 'da': 'Pad See Ew'},
 'opt2': {'en': 'Pad Thai', 'de': 'Pad Thai', 'es': 'Pad Thai', 'nl': 'Pad Thai', 'sv': 'Pad Thai', 'fr': 'Pad thai', 'da': 'Pad Thai'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which body system includes the stomach and intestines?', 'de': 'Welches Koerpersystem umfasst den Magen und den Darm?', 'es': '¿Que sistema corporal incluye el estomago y los intestinos?', 'nl': 'Welk lichaamssysteem omvat de maag en de darmen?', 'sv': 'Vilket kroppssystem inkluderar magen och tarmarna?', 'fr': "Quel systeme corporel comprend l'estomac et les intestins ?", 'da': 'Hvilket kropssystem omfatter mavesækken og tarmene?'},
 'opt1': {'en': 'Digestive system', 'de': 'Verdauungssystem', 'es': 'Sistema digestivo', 'nl': 'Spijsverteringsstelsel', 'sv': 'Matsmaltningssystemet', 'fr': 'Le systeme digestif', 'da': 'Fordøjelsessystemet'},
 'opt2': {'en': 'Circulatory system', 'de': 'Kreislaufsystem', 'es': 'Sistema circulatorio', 'nl': 'Bloedsomloop', 'sv': 'Cirkulationssystemet', 'fr': 'Le systeme circulatoire', 'da': 'Kredsløbssystemet'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Carthage?', 'de': 'In welchem Land liegt die antike Stadt Karthago?', 'es': '¿En que pais esta la antigua ciudad de Cartago?', 'nl': 'In welk land ligt de oude stad Carthago?', 'sv': 'I vilket land ligger den antika staden Karthago?', 'fr': "Dans quel pays se trouve l'ancienne cite de Carthage ?", 'da': 'Hvilket land huser den antikke by Karthago?'},
 'opt1': {'en': 'Morocco', 'de': 'Marokko', 'es': 'Marruecos', 'nl': 'Marokko', 'sv': 'Marocko', 'fr': 'le Maroc', 'da': 'Marokko'},
 'opt2': {'en': 'Tunisia', 'de': 'Tunesien', 'es': 'Tunez', 'nl': 'Tunesie', 'sv': 'Tunisien', 'fr': 'la Tunisie', 'da': 'Tunesien'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which First Lady was the first to graduate from an all-women's college before entering the White House?", 'de': 'Welche First Lady war die erste, die vor dem Einzug ins Weisse Haus ein reines Frauencollege abschloss?', 'es': '¿Que primera dama fue la primera en graduarse de una universidad solo para mujeres antes de entrar a la Casa Blanca?', 'nl': 'Welke First Lady was de eerste die afstudeerde aan een vrouwencollege voordat ze het Witte Huis binnentrad?', 'sv': 'Vilken First Lady var forst med att ta examen fran en rent kvinnlig hogskola innan hon flyttade in i Vita huset?', 'fr': "Quelle Première dame a ete la premiere a etre diplomee d'un college reserve aux femmes avant d'entrer a la Maison-Blanche ?", 'da': 'Hvilken førstedame var den første til at dimittere fra et rent kvindecollege, før hun kom til Det Hvide Hus?'},
 'opt1': {'en': 'Hillary Clinton', 'de': 'Hillary Clinton', 'es': 'Hillary Clinton', 'nl': 'Hillary Clinton', 'sv': 'Hillary Clinton', 'fr': 'Hillary Clinton', 'da': 'Hillary Clinton'},
 'opt2': {'en': 'Barbara Bush', 'de': 'Barbara Bush', 'es': 'Barbara Bush', 'nl': 'Barbara Bush', 'sv': 'Barbara Bush', 'fr': 'Barbara Bush', 'da': 'Barbara Bush'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which composer wrote 'The Four Seasons' concertos?", 'de': "Welcher Komponist schrieb die Konzerte 'Die vier Jahreszeiten'?", 'es': "¿Que compositor escribio los conciertos de 'Las cuatro estaciones'?", 'nl': "Welke componist schreef de concerten 'De Vier Jaargetijden'?", 'sv': "Vilken kompositor skrev konserterna 'De fyra arstiderna'?", 'fr': "Quel compositeur a ecrit les concertos des 'Quatre Saisons' ?", 'da': "Hvilken komponist skrev koncerterne 'De fire årstider'?"},
 'opt1': {'en': 'Johann Sebastian Bach', 'de': 'Johann Sebastian Bach', 'es': 'Johann Sebastian Bach', 'nl': 'Johann Sebastian Bach', 'sv': 'Johann Sebastian Bach', 'fr': 'Johann Sebastian Bach', 'da': 'Johann Sebastian Bach'},
 'opt2': {'en': 'Antonio Vivaldi', 'de': 'Antonio Vivaldi', 'es': 'Antonio Vivaldi', 'nl': 'Antonio Vivaldi', 'sv': 'Antonio Vivaldi', 'fr': 'Antonio Vivaldi', 'da': 'Antonio Vivaldi'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which actress played Hermione Granger in the Harry Potter film series?', 'de': 'Welche Schauspielerin spielte Hermine Granger in der Harry-Potter-Filmreihe?', 'es': '¿Que actriz interpreto a Hermione Granger en la serie de peliculas de Harry Potter?', 'nl': 'Welke actrice speelde Hermelien Griffel in de Harry Potter-filmreeks?', 'sv': 'Vilken skadespelerska spelade Hermione Granger i Harry Potter-filmerna?', 'fr': "Quelle actrice a joue Hermione Granger dans la serie de films Harry Potter ?", 'da': 'Hvilken skuespiller spillede Hermione Granger i Harry Potter-filmserien?'},
 'opt1': {'en': 'Emma Watson', 'de': 'Emma Watson', 'es': 'Emma Watson', 'nl': 'Emma Watson', 'sv': 'Emma Watson', 'fr': 'Emma Watson', 'da': 'Emma Watson'},
 'opt2': {'en': 'Emma Roberts', 'de': 'Emma Roberts', 'es': 'Emma Roberts', 'nl': 'Emma Roberts', 'sv': 'Emma Roberts', 'fr': 'Emma Roberts', 'da': 'Emma Roberts'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Vietnamese soup is made with broth, rice noodles, and herbs?', 'de': 'Welche vietnamesische Suppe wird mit Bruehe, Reisnudeln und Kraeutern zubereitet?', 'es': '¿Que sopa vietnamita se hace con caldo, fideos de arroz y hierbas?', 'nl': 'Welke Vietnamese soep wordt gemaakt met bouillon, rijstnoedels en kruiden?', 'sv': 'Vilken vietnamesisk soppa gors med buljong, risnudlar och ortor?', 'fr': "Quelle soupe vietnamienne est faite de bouillon, de nouilles de riz et d'herbes ?", 'da': 'Hvilken vietnamesisk suppe er lavet med bouillon, risnudler og krydderurter?'},
 'opt1': {'en': 'Bun bo Hue', 'de': 'Bun Bo Hue', 'es': 'Bun bo Hue', 'nl': 'Bun bo Hue', 'sv': 'Bun bo Hue', 'fr': 'Bun bo Hue', 'da': 'Bun bo Hue'},
 'opt2': {'en': 'Pho', 'de': 'Pho', 'es': 'Pho', 'nl': 'Pho', 'sv': 'Pho', 'fr': 'Pho', 'da': 'Pho'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Outer Banks?', 'de': 'In welchem US-Bundesstaat liegen die Outer Banks?', 'es': '¿En que estado de EE.UU. estan los Outer Banks?', 'nl': 'In welke Amerikaanse staat liggen de Outer Banks?', 'sv': 'I vilken amerikansk delstat ligger Outer Banks?', 'fr': "Dans quel État americain se trouvent les Outer Banks ?", 'da': 'I hvilken amerikansk delstat ligger Outer Banks?'},
 'opt1': {'en': 'North Carolina', 'de': 'North Carolina', 'es': 'Carolina del Norte', 'nl': 'North Carolina', 'sv': 'North Carolina', 'fr': 'la Caroline du Nord', 'da': 'North Carolina'},
 'opt2': {'en': 'Virginia', 'de': 'Virginia', 'es': 'Virginia', 'nl': 'Virginia', 'sv': 'Virginia', 'fr': 'la Virginie', 'da': 'Virginia'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which small mammal is known for its quills used as defense?', 'de': 'Welches kleine Saeugetier ist fuer seine Stacheln bekannt, die zur Verteidigung dienen?', 'es': '¿Que pequeño mamifero es conocido por sus puas usadas como defensa?', 'nl': 'Welk klein zoogdier staat bekend om zijn stekels die als verdediging worden gebruikt?', 'sv': 'Vilket litet daggdjur ar kant for sina taggar som anvands som forsvar?', 'fr': "Quel petit mammifere est connu pour ses piquants utilises en defense ?", 'da': 'Hvilket lille pattedyr er kendt for sine pigge brugt til forsvar?'},
 'opt1': {'en': 'Hedgehog', 'de': 'Igel', 'es': 'Erizo', 'nl': 'Egel', 'sv': 'Igelkott', 'fr': 'Herisson', 'da': 'Pindsvin'},
 'opt2': {'en': 'Porcupine', 'de': 'Stachelschwein', 'es': 'Puerco espin', 'nl': 'Stekelvarken', 'sv': 'Piggsvin', 'fr': 'Porc-epic', 'da': 'Pindsvin'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which Mexican holiday celebrates a girl's 15th birthday with a large party?", 'de': "Welcher mexikanische Feiertag feiert den 15. Geburtstag eines Maedchens mit einer grossen Party?", 'es': "¿Que festividad mexicana celebra el decimoquinto cumpleaños de una niña con una gran fiesta?", 'nl': "Welke Mexicaanse feestdag viert de 15e verjaardag van een meisje met een groot feest?", 'sv': "Vilken mexikansk hogtid firar en flickas 15-arsdag med en stor fest?", 'fr': "Quelle fete mexicaine celebre le 15e anniversaire d'une jeune fille avec une grande fete ?", 'da': 'Hvilken mexicansk helligdag fejrer en piges 15-års fødselsdag med en stor fest?'},
 'opt1': {'en': 'Quinceanera', 'de': 'Quinceañera', 'es': 'Quinceañera', 'nl': 'Quinceañera', 'sv': 'Quinceañera', 'fr': 'La quinceañera', 'da': 'Quinceañera'},
 'opt2': {'en': 'Dia de los Ninos', 'de': 'Dia de los Niños', 'es': 'Dia de los Niños', 'nl': 'Dia de los Niños', 'sv': 'Dia de los Niños', 'fr': 'Dia de los Niños', 'da': 'Dia de los Niños'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'The Bell Jar'?", 'de': "Welche Autorin schrieb 'Die Glasglocke'?", 'es': "¿Que autora escribio 'La campana de cristal'?", 'nl': "Welke auteur schreef 'De glazen stolp'?", 'sv': "Vilken forfattare skrev 'Glaskupan'?", 'fr': "Quelle autrice a ecrit 'La Cloche de detresse' ?", 'da': "Hvilken forfatter skrev 'Glasklokken'?"},
 'opt1': {'en': 'Anne Sexton', 'de': 'Anne Sexton', 'es': 'Anne Sexton', 'nl': 'Anne Sexton', 'sv': 'Anne Sexton', 'fr': 'Anne Sexton', 'da': 'Anne Sexton'},
 'opt2': {'en': 'Sylvia Plath', 'de': 'Sylvia Plath', 'es': 'Sylvia Plath', 'nl': 'Sylvia Plath', 'sv': 'Sylvia Plath', 'fr': 'Sylvia Plath', 'da': 'Sylvia Plath'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient ruins of Ephesus and the city of Istanbul?', 'de': 'In welchem Land liegen die antiken Ruinen von Ephesus und die Stadt Istanbul?', 'es': '¿En que pais estan las ruinas antiguas de Efeso y la ciudad de Estambul?', 'nl': 'In welk land liggen de oude ruines van Efeze en de stad Istanbul?', 'sv': 'I vilket land ligger de antika ruinerna av Efesos och staden Istanbul?', 'fr': "Dans quel pays se trouvent les ruines antiques d'Ephese et la ville d'Istanbul ?", 'da': 'Hvilket land huser de gamle ruiner af Efesos og byen Istanbul?'},
 'opt1': {'en': 'Turkey', 'de': 'Tuerkei', 'es': 'Turquia', 'nl': 'Turkije', 'sv': 'Turkiet', 'fr': 'la Turquie', 'da': 'Tyrkiet'},
 'opt2': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which instrument is central to Andean folk music, similar to a small guitar?', 'de': 'Welches Instrument ist zentral fuer die andine Volksmusik, aehnlich einer kleinen Gitarre?', 'es': '¿Que instrumento es central en la musica folclorica andina, similar a una guitarra pequeña?', 'nl': 'Welk instrument is centraal in Andesvolksmuziek, vergelijkbaar met een kleine gitaar?', 'sv': 'Vilket instrument ar centralt i andinsk folkmusik, liknande en liten gitarr?', 'fr': "Quel instrument est central dans la musique folklorique andine, semblable a une petite guitare ?", 'da': 'Hvilket instrument er centralt i andinsk folkemusik, ligner en lille guitar?'},
 'opt1': {'en': 'Charango', 'de': 'Charango', 'es': 'Charango', 'nl': 'Charango', 'sv': 'Charango', 'fr': 'Charango', 'da': 'Charango'},
 'opt2': {'en': 'Ukulele', 'de': 'Ukulele', 'es': 'Ukelele', 'nl': 'Ukelele', 'sv': 'Ukulele', 'fr': 'Ukulele', 'da': 'Ukulele'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dessert is a deep-fried pastry often filled with fruit?', 'de': 'Welches suddstaatliche Dessert ist ein frittiertes Gebaeck, oft mit Obst gefuellt?', 'es': '¿Que postre sureño de EE.UU. es una masa frita a menudo rellena de fruta?', 'nl': 'Welk zuidelijk Amerikaans dessert is een gefrituurd gebak vaak gevuld met fruit?', 'sv': 'Vilken sydstatlig dessert ar ett friterat bakverk ofta fyllt med frukt?', 'fr': "Quel dessert du Sud des États-Unis est une patisserie frite souvent fourree de fruits ?", 'da': 'Hvilken sydstatsdessert er en friturestegt bagværk ofte fyldt med frugt?'},
 'opt1': {'en': 'Fried pie', 'de': 'Frittierte Pastete', 'es': 'Pastel frito', 'nl': 'Gefrituurde taart', 'sv': 'Friterad pai', 'fr': 'Tarte frite', 'da': 'Friturestegt tærte'},
 'opt2': {'en': 'Cobbler', 'de': 'Cobbler', 'es': 'Cobbler', 'nl': 'Cobbler', 'sv': 'Cobbler', 'fr': 'Cobbler', 'da': 'Cobbler'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which mineral, found in bananas, helps regulate blood pressure and muscle function?', 'de': 'Welches in Bananen enthaltene Mineral hilft, den Blutdruck und die Muskelfunktion zu regulieren?', 'es': '¿Que mineral, presente en los platanos, ayuda a regular la presion arterial y la funcion muscular?', 'nl': 'Welk mineraal, aanwezig in bananen, helpt de bloeddruk en spierfunctie te reguleren?', 'sv': 'Vilket mineral, som finns i bananer, hjalper till att reglera blodtryck och muskelfunktion?', 'fr': "Quel mineral, present dans les bananes, aide a reguler la pression arterielle et la fonction musculaire ?", 'da': 'Hvilket mineral, fundet i bananer, hjælper med at regulere blodtryk og muskelfunktion?'},
 'opt1': {'en': 'Potassium', 'de': 'Kalium', 'es': 'Potasio', 'nl': 'Kalium', 'sv': 'Kalium', 'fr': 'Potassium', 'da': 'Kalium'},
 'opt2': {'en': 'Magnesium', 'de': 'Magnesium', 'es': 'Magnesio', 'nl': 'Magnesium', 'sv': 'Magnesium', 'fr': 'Magnesium', 'da': 'Magnesium'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which film has Meryl Streep playing a fashion magazine editor?', 'de': 'In welchem Film spielt Meryl Streep eine Modemagazin-Redakteurin?', 'es': '¿En que pelicula Meryl Streep interpreta a una editora de una revista de moda?', 'nl': 'In welke film speelt Meryl Streep een hoofdredacteur van een modetijdschrift?', 'sv': 'I vilken film spelar Meryl Streep en modetidningsredaktor?', 'fr': "Dans quel film Meryl Streep joue-t-elle une redactrice de magazine de mode ?", 'da': 'I hvilken film spiller Meryl Streep en modemagasin-redaktør?'},
 'opt1': {'en': 'The Post', 'de': 'Die Verlegerin', 'es': 'Los archivos del Pentagono', 'nl': 'The Post', 'sv': 'The Post', 'fr': 'Pentagon Papers', 'da': 'The Post'},
 'opt2': {'en': 'The Devil Wears Prada', 'de': 'Der Teufel traegt Prada', 'es': 'El diablo viste a la moda', 'nl': 'The Devil Wears Prada', 'sv': 'Djavulen bar Prada', 'fr': "Le Diable s'habille en Prada", 'da': 'Djævelen går i Prada'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient ruins of Palmyra?', 'de': 'In welchem Land liegen die antiken Ruinen von Palmyra?', 'es': '¿En que pais estan las ruinas antiguas de Palmira?', 'nl': 'In welk land liggen de oude ruines van Palmyra?', 'sv': 'I vilket land ligger de antika ruinerna av Palmyra?', 'fr': "Dans quel pays se trouvent les ruines antiques de Palmyre ?", 'da': 'Hvilket land huser de gamle ruiner af Palmyra?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Syria', 'de': 'Syrien', 'es': 'Siria', 'nl': 'Syrie', 'sv': 'Syrien', 'fr': 'la Syrie', 'da': 'Syrien'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which type of coat, named after a British field marshal, has a belted waist and epaulettes?', 'de': 'Welche Art von Mantel, benannt nach einem britischen Feldmarschall, hat eine Guertel-Taille und Schulterklappen?', 'es': '¿Que tipo de abrigo, nombrado en honor a un mariscal de campo britanico, tiene cintura con cinturon y hombreras?', 'nl': 'Welk type jas, vernoemd naar een Britse veldmaarschalk, heeft een riemtaille en epauletten?', 'sv': 'Vilken typ av kappa, uppkallad efter en brittisk faltmarskalk, har en bältestalje och axelklaffar?', 'fr': "Quel type de manteau, nomme d'apres un marechal britannique, a une taille cintree et des epaulettes ?", 'da': 'Hvilken type frakke, opkaldt efter en britisk feltmarskal, har en bæltetalje og epauletter?'},
 'opt1': {'en': 'Duffle coat', 'de': 'Dufflecoat', 'es': 'Trenca', 'nl': 'Duffelcoat', 'sv': 'Dufflecoat', 'fr': 'Duffle-coat', 'da': 'Duffelfrakke'},
 'opt2': {'en': 'Trench coat', 'de': 'Trenchcoat', 'es': 'Gabardina', 'nl': 'Trenchcoat', 'sv': 'Trenchcoat', 'fr': 'Trench-coat', 'da': 'Trenchcoat'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large flightless bird native to Africa can run up to 43 mph?', 'de': 'Welcher grosse flugunfaehige Vogel aus Afrika kann bis zu 70 km/h laufen?', 'es': '¿Que gran ave no voladora nativa de Africa puede correr hasta 70 km/h?', 'nl': 'Welke grote vliegloze vogel, inheems in Afrika, kan tot 70 km/u rennen?', 'sv': 'Vilken stor flygofor fagel, inhemsk i Afrika, kan springa upp till 70 km/h?', 'fr': "Quel grand oiseau incapable de voler, originaire d'Afrique, peut courir jusqu'a 70 km/h ?", 'da': 'Hvilken stor flyveudygtig fugl, hjemmehørende i Afrika, kan løbe op til 70 km/t?'},
 'opt1': {'en': 'Ostrich', 'de': 'Strauss', 'es': 'Avestruz', 'nl': 'Struisvogel', 'sv': 'Struts', 'fr': 'Autruche', 'da': 'Struds'},
 'opt2': {'en': 'Emu', 'de': 'Emu', 'es': 'Emu', 'nl': 'Emoe', 'sv': 'Emu', 'fr': 'Emeu', 'da': 'Emu'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which First Lady was the first to be a working mother while her husband was in office?', 'de': 'Welche First Lady war die erste berufstaetige Mutter, waehrend ihr Mann im Amt war?', 'es': '¿Que primera dama fue la primera madre trabajadora mientras su esposo estaba en el cargo?', 'nl': 'Welke First Lady was de eerste werkende moeder terwijl haar man in functie was?', 'sv': 'Vilken First Lady var forst med att vara en yrkesarbetande mor medan hennes man var i amBetet?', 'fr': "Quelle Première dame a ete la premiere mere active professionnellement pendant que son mari etait en fonction ?", 'da': 'Hvilken førstedame var den første til at være en erhvervsarbejdende mor, mens hendes mand var i embedet?'},
 'opt1': {'en': 'Rosalynn Carter', 'de': 'Rosalynn Carter', 'es': 'Rosalynn Carter', 'nl': 'Rosalynn Carter', 'sv': 'Rosalynn Carter', 'fr': 'Rosalynn Carter', 'da': 'Rosalynn Carter'},
 'opt2': {'en': 'Hillary Clinton', 'de': 'Hillary Clinton', 'es': 'Hillary Clinton', 'nl': 'Hillary Clinton', 'sv': 'Hillary Clinton', 'fr': 'Hillary Clinton', 'da': 'Hillary Clinton'}},

{'category': 'History', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which First Lady established the White House's first permanent library?", 'de': 'Welche First Lady richtete die erste dauerhafte Bibliothek des Weissen Hauses ein?', 'es': '¿Que primera dama establecio la primera biblioteca permanente de la Casa Blanca?', 'nl': 'Welke First Lady richtte de eerste permanente bibliotheek van het Witte Huis op?', 'sv': 'Vilken First Lady etablerade Vita husets forsta permanenta bibliotek?', 'fr': "Quelle Première dame a etabli la premiere bibliotheque permanente de la Maison-Blanche ?", 'da': 'Hvilken førstedame etablerede Det Hvide Hus’ første permanente bibliotek?'},
 'opt1': {'en': 'Lou Hoover', 'de': 'Lou Hoover', 'es': 'Lou Hoover', 'nl': 'Lou Hoover', 'sv': 'Lou Hoover', 'fr': 'Lou Hoover', 'da': 'Lou Hoover'},
 'opt2': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Chilean poet won the Nobel Prize in Literature in 1971?', 'de': 'Welcher chilenische Dichter gewann 1971 den Nobelpreis fuer Literatur?', 'es': '¿Que poeta chileno gano el Premio Nobel de Literatura en 1971?', 'nl': 'Welke Chileense dichter won in 1971 de Nobelprijs voor Literatuur?', 'sv': 'Vilken chilensk poet vann Nobelpriset i litteratur 1971?', 'fr': "Quel poete chilien a remporte le prix Nobel de litterature en 1971 ?", 'da': 'Hvilken chilensk digter vandt Nobelprisen i litteratur i 1971?'},
 'opt1': {'en': 'Pablo Neruda', 'de': 'Pablo Neruda', 'es': 'Pablo Neruda', 'nl': 'Pablo Neruda', 'sv': 'Pablo Neruda', 'fr': 'Pablo Neruda', 'da': 'Pablo Neruda'},
 'opt2': {'en': 'Octavio Paz', 'de': 'Octavio Paz', 'es': 'Octavio Paz', 'nl': 'Octavio Paz', 'sv': 'Octavio Paz', 'fr': 'Octavio Paz', 'da': 'Octavio Paz'}},

{'category': 'Science', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which vitamin is essential for night vision and found in carrots?', 'de': 'Welches Vitamin ist essentiell fuer die Nachtsicht und in Karotten enthalten?', 'es': '¿Que vitamina es esencial para la vision nocturna y se encuentra en las zanahorias?', 'nl': 'Welk vitamine is essentieel voor nachtzicht en te vinden in wortels?', 'sv': 'Vilket vitamin ar viktigt for nattseende och finns i morotter?', 'fr': "Quelle vitamine est essentielle a la vision nocturne et se trouve dans les carottes ?", 'da': 'Hvilket vitamin er essentielt for natsyn og findes i gulerødder?'},
 'opt1': {'en': 'Vitamin A', 'de': 'Vitamin A', 'es': 'Vitamina A', 'nl': 'Vitamine A', 'sv': 'Vitamin A', 'fr': 'Vitamine A', 'da': 'Vitamin A'},
 'opt2': {'en': 'Vitamin C', 'de': 'Vitamin C', 'es': 'Vitamina C', 'nl': 'Vitamine C', 'sv': 'Vitamin C', 'fr': 'Vitamine C', 'da': 'Vitamin C'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which Austrian painter is known for 'The Kiss', featuring gold leaf patterns?", 'de': "Welcher oesterreichische Maler ist fuer 'Der Kuss' bekannt, mit Goldblattmustern?", 'es': "¿Que pintor austriaco es conocido por 'El beso', con patrones de pan de oro?", 'nl': "Welke Oostenrijkse schilder staat bekend om 'De Kus', met bladgoudpatronen?", 'sv': "Vilken osterrikisk malare ar kand for 'Kyssen', med bladguldmonster?", 'fr': "Quel peintre autrichien est connu pour 'Le Baiser', avec ses motifs en feuille d'or ?", 'da': "Hvilken østrigsk maler er kendt for 'Kysset', med bladguldmønstre?"},
 'opt1': {'en': 'Egon Schiele', 'de': 'Egon Schiele', 'es': 'Egon Schiele', 'nl': 'Egon Schiele', 'sv': 'Egon Schiele', 'fr': 'Egon Schiele', 'da': 'Egon Schiele'},
 'opt2': {'en': 'Gustav Klimt', 'de': 'Gustav Klimt', 'es': 'Gustav Klimt', 'nl': 'Gustav Klimt', 'sv': 'Gustav Klimt', 'fr': 'Gustav Klimt', 'da': 'Gustav Klimt'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which Himalayan country is known as the 'Land of the Thunder Dragon'?", 'de': "Welches Himalaya-Land ist als 'Land des Donnerdrachen' bekannt?", 'es': "¿Que pais himalayo es conocido como la 'Tierra del Dragon del Trueno'?", 'nl': "Welk Himalaya-land staat bekend als het 'Land van de Donderdraak'?", 'sv': "Vilket Himalayaland ar kant som 'Landet dar askans drake bor'?", 'fr': "Quel pays de l'Himalaya est connu comme le 'Pays du dragon tonnerre' ?", 'da': "Hvilket Himalaya-land er kendt som 'Tordendragens land'?"},
 'opt1': {'en': 'Nepal', 'de': 'Nepal', 'es': 'Nepal', 'nl': 'Nepal', 'sv': 'Nepal', 'fr': 'le Nepal', 'da': 'Nepal'},
 'opt2': {'en': 'Bhutan', 'de': 'Bhutan', 'es': 'Butan', 'nl': 'Bhutan', 'sv': 'Bhutan', 'fr': 'le Bhoutan', 'da': 'Bhutan'}},

{'category': 'Music', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which German composer wrote the six 'Brandenburg Concertos'?", 'de': "Welcher deutsche Komponist schrieb die sechs 'Brandenburgischen Konzerte'?", 'es': "¿Que compositor aleman escribio los seis 'Conciertos de Brandeburgo'?", 'nl': "Welke Duitse componist schreef de zes 'Brandenburgse Concerten'?", 'sv': "Vilken tysk kompositor skrev de sex 'Brandenburgkonserterna'?", 'fr': "Quel compositeur allemand a ecrit les six 'Concertos brandebourgeois' ?", 'da': "Hvilken tysk komponist skrev de seks 'Brandenburgkoncerter'?"},
 'opt1': {'en': 'Georg Friedrich Handel', 'de': 'Georg Friedrich Haendel', 'es': 'Georg Friedrich Handel', 'nl': 'Georg Friedrich Handel', 'sv': 'Georg Friedrich Handel', 'fr': 'Georg Friedrich Haendel', 'da': 'Georg Friedrich Handel'},
 'opt2': {'en': 'Johann Sebastian Bach', 'de': 'Johann Sebastian Bach', 'es': 'Johann Sebastian Bach', 'nl': 'Johann Sebastian Bach', 'sv': 'Johann Sebastian Bach', 'fr': 'Johann Sebastian Bach', 'da': 'Johann Sebastian Bach'}},

{'category': 'Food', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which Hungarian dish is a paprika-spiced beef stew?', 'de': 'Welches ungarische Gericht ist ein mit Paprika gewuerzter Rindereintopf?', 'es': '¿Que plato hungaro es un guiso de carne condimentado con pimenton?', 'nl': 'Welk Hongaars gerecht is een met paprika gekruide rundvleesstoofpot?', 'sv': 'Vilken ungersk ratt ar en paprikakryddad nautgryta?', 'fr': "Quel plat hongrois est un ragout de boeuf epice au paprika ?", 'da': 'Hvilken ungarsk ret er en paprikakrydret oksekødsgryde?'},
 'opt1': {'en': 'Paprikash', 'de': 'Paprikasch', 'es': 'Paprikash', 'nl': 'Paprikash', 'sv': 'Paprikas', 'fr': 'Paprikas', 'da': 'Paprikas'},
 'opt2': {'en': 'Goulash', 'de': 'Gulasch', 'es': 'Gulasch', 'nl': 'Goulash', 'sv': 'Gulasch', 'fr': 'Goulasch', 'da': 'Gullasch'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Hindu mythology, which god is depicted with an elephant head?', 'de': 'In der hinduistischen Mythologie, welcher Gott wird mit einem Elefantenkopf dargestellt?', 'es': '¿En la mitologia hindu, que dios se representa con cabeza de elefante?', 'nl': 'In de hindoemythologie, welke god wordt afgebeeld met een olifantenkop?', 'sv': 'I hinduisk mytologi, vilken gud avbildas med ett elefanthuvud?', 'fr': "Dans la mythologie hindoue, quel dieu est represente avec une tete d'elephant ?", 'da': 'I hinduistisk mytologi, hvilken gud er afbildet med et elefanthoved?'},
 'opt1': {'en': 'Ganesha', 'de': 'Ganesha', 'es': 'Ganesha', 'nl': 'Ganesha', 'sv': 'Ganesha', 'fr': 'Ganesha', 'da': 'Ganesha'},
 'opt2': {'en': 'Hanuman', 'de': 'Hanuman', 'es': 'Hanuman', 'nl': 'Hanuman', 'sv': 'Hanuman', 'fr': 'Hanuman', 'da': 'Hanuman'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which actress won an Oscar for playing a woman with dissociative identity disorder in 'Sybil'?", 'de': "Welche Schauspielerin gewann einen Oscar fuer ihre Rolle als Frau mit dissoziativer Identitaetsstoerung in 'Sybil'?", 'es': "¿Que actriz gano un Oscar por interpretar a una mujer con trastorno de identidad disociativo en 'Sybil'?", 'nl': "Welke actrice won een Oscar voor haar rol als vrouw met dissociatieve identiteitsstoornis in 'Sybil'?", 'sv': "Vilken skadespelerska vann en Oscar for att ha spelat en kvinna med dissociativ identitetsstorning i 'Sybil'?", 'fr': "Quelle actrice a remporte un Oscar pour son role d'une femme atteinte de trouble dissociatif de l'identite dans 'Sybil' ?", 'da': "Hvilken skuespiller vandt en Oscar for at spille en kvinde med dissociativ identitetsforstyrrelse i 'Sybil'?"},
 'opt1': {'en': 'Sally Field', 'de': 'Sally Field', 'es': 'Sally Field', 'nl': 'Sally Field', 'sv': 'Sally Field', 'fr': 'Sally Field', 'da': 'Sally Field'},
 'opt2': {'en': 'Jane Fonda', 'de': 'Jane Fonda', 'es': 'Jane Fonda', 'nl': 'Jane Fonda', 'sv': 'Jane Fonda', 'fr': 'Jane Fonda', 'da': 'Jane Fonda'}},

{'category': 'Space', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which moon of Saturn is known for its geysers of water ice?', 'de': 'Welcher Saturnmond ist fuer seine Geysire aus Wassereis bekannt?', 'es': '¿Que luna de Saturno es conocida por sus geiseres de hielo de agua?', 'nl': 'Welke maan van Saturnus staat bekend om zijn geisers van waterijs?', 'sv': 'Vilken av Saturnus manar ar kand for sina geysrar av isvatten?', 'fr': "Quelle lune de Saturne est connue pour ses geysers de glace d'eau ?", 'da': 'Hvilken af Saturns måner er kendt for sine gejsere af vandis?'},
 'opt1': {'en': 'Enceladus', 'de': 'Enceladus', 'es': 'Encelado', 'nl': 'Enceladus', 'sv': 'Enceladus', 'fr': 'Encelade', 'da': 'Enceladus'},
 'opt2': {'en': 'Titan', 'de': 'Titan', 'es': 'Titan', 'nl': 'Titan', 'sv': 'Titan', 'fr': 'Titan', 'da': 'Titan'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which African country was formerly known as the Gold Coast before independence in 1957?', 'de': 'Welches afrikanische Land wurde vor der Unabhaengigkeit 1957 als Goldkueste bezeichnet?', 'es': '¿Que pais africano se conocia anteriormente como Costa de Oro antes de la independencia en 1957?', 'nl': 'Welk Afrikaans land stond voorheen bekend als de Goudkust voor onafhankelijkheid in 1957?', 'sv': 'Vilket afrikanskt land var tidigare kant som Guldkusten fore sjalvstandigheten 1957?', 'fr': "Quel pays africain etait autrefois connu sous le nom de Cote-de-l'Or avant l'independance en 1957 ?", 'da': 'Hvilket afrikansk land var tidligere kendt som Guldkysten før uafhængigheden i 1957?'},
 'opt1': {'en': 'Ghana', 'de': 'Ghana', 'es': 'Ghana', 'nl': 'Ghana', 'sv': 'Ghana', 'fr': 'le Ghana', 'da': 'Ghana'},
 'opt2': {'en': 'Nigeria', 'de': 'Nigeria', 'es': 'Nigeria', 'nl': 'Nigeria', 'sv': 'Nigeria', 'fr': 'le Nigeria', 'da': 'Nigeria'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which French designer founded a house known for its quilted handbags and tweed suits?', 'de': 'Welche franzoesische Designerin gruendete ein Haus, das fuer seine gesteppten Handtaschen und Tweed-Anzuege bekannt ist?', 'es': '¿Que diseñadora francesa fundo una casa conocida por sus bolsos acolchados y trajes de tweed?', 'nl': 'Welke Franse ontwerper richtte een huis op dat bekend staat om zijn gewatteerde handtassen en tweed pakken?', 'sv': 'Vilken fransk designer grundade ett hus kant for sina quiltade handvaskor och tweeddrakter?', 'fr': "Quelle creatrice française a fonde une maison connue pour ses sacs a main matelasses et ses tailleurs en tweed ?", 'da': 'Hvilken fransk designer grundlagde et hus kendt for sine quiltede håndtasker og tweed-jakkesæt?'},
 'opt1': {'en': 'Yves Saint Laurent', 'de': 'Yves Saint Laurent', 'es': 'Yves Saint Laurent', 'nl': 'Yves Saint Laurent', 'sv': 'Yves Saint Laurent', 'fr': 'Yves Saint Laurent', 'da': 'Yves Saint Laurent'},
 'opt2': {'en': 'Coco Chanel', 'de': 'Coco Chanel', 'es': 'Coco Chanel', 'nl': 'Coco Chanel', 'sv': 'Coco Chanel', 'fr': 'Coco Chanel', 'da': 'Coco Chanel'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Danish philosopher is considered the father of existentialism?', 'de': 'Welcher daenische Philosoph gilt als Vater des Existentialismus?', 'es': '¿Que filosofo danes es considerado el padre del existencialismo?', 'nl': 'Welke Deense filosoof wordt beschouwd als de vader van het existentialisme?', 'sv': 'Vilken dansk filosof anses vara existentialismens fader?', 'fr': "Quel philosophe danois est considere comme le pere de l'existentialisme ?", 'da': 'Hvilken dansk filosof anses for eksistentialismens fader?'},
 'opt1': {'en': 'Soren Kierkegaard', 'de': 'Soeren Kierkegaard', 'es': 'Soren Kierkegaard', 'nl': 'Soren Kierkegaard', 'sv': 'Soren Kierkegaard', 'fr': 'Soren Kierkegaard', 'da': 'Søren Kierkegaard'},
 'opt2': {'en': 'Friedrich Nietzsche', 'de': 'Friedrich Nietzsche', 'es': 'Friedrich Nietzsche', 'nl': 'Friedrich Nietzsche', 'sv': 'Friedrich Nietzsche', 'fr': 'Friedrich Nietzsche', 'da': 'Friedrich Nietzsche'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which quilting pattern uses triangles arranged to look like flying geese?', 'de': 'Welches Quiltmuster verwendet Dreiecke, die wie fliegende Gaense angeordnet sind?', 'es': '¿Que patron de acolchado usa triangulos dispuestos para parecer gansos volando?', 'nl': 'Welk quiltpatroon gebruikt driehoeken gerangschikt om op vliegende ganzen te lijken?', 'sv': 'Vilket lapptacksmonster anvander trianglar arrangerade for att se ut som flygande gass?', 'fr': "Quel motif de patchwork utilise des triangles arranges pour ressembler a des oies volantes ?", 'da': 'Hvilket quiltmønster bruger trekanter arrangeret til at ligne flyvende gæs?'},
 'opt1': {'en': 'Flying Geese', 'de': 'Fliegende Gaense', 'es': 'Gansos voladores', 'nl': 'Vliegende Ganzen', 'sv': 'Flygande gass', 'fr': 'Oies volantes', 'da': 'Flyvende Gæs'},
 'opt2': {'en': "Bear's Paw", 'de': 'Baerentatze', 'es': 'Pata de oso', 'nl': 'Berenklauw', 'sv': 'Bjornlabb', 'fr': "Patte d'ours", 'da': 'Bjørnepote'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which desert succulent blooms once every several decades before dying, native to the Americas?', 'de': 'Welche Wuestensukkulente aus Amerika bluht nur einmal alle paar Jahrzehnte, bevor sie abstirbt?', 'es': '¿Que suculenta del desierto, nativa de America, florece una vez cada varias decadas antes de morir?', 'nl': 'Welke woestijnvetplant, inheems in Amerika, bloeit slechts eens per meerdere decennia voordat hij sterft?', 'sv': 'Vilken oken-suckulent, inhemsk i Amerika, blommar bara en gang vart flera artionde innan den dor?', 'fr': "Quelle plante grasse du desert, originaire des Ameriques, fleurit une fois tous les plusieurs decennies avant de mourir ?", 'da': 'Hvilken ørken-sukkulent, hjemmehørende i Amerika, blomstrer kun én gang hvert par årtier, før den dør?'},
 'opt1': {'en': 'Century plant (Agave)', 'de': 'Jahrhundertpflanze (Agave)', 'es': 'Planta del siglo (Agave)', 'nl': 'Eeuwplant (Agave)', 'sv': 'Arhundradeplanta (Agave)', 'fr': 'Agave americaine', 'da': 'Århundredeplante (Agave)'},
 'opt2': {'en': 'Saguaro cactus', 'de': 'Saguaro-Kaktus', 'es': 'Cactus saguaro', 'nl': 'Saguarocactus', 'sv': 'Saguarokaktus', 'fr': 'Cactus saguaro', 'da': 'Saguarokaktus'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which architect designed the Fallingwater house built over a waterfall?', 'de': 'Welcher Architekt entwarf das Haus Fallingwater, das ueber einem Wasserfall errichtet wurde?', 'es': '¿Que arquitecto diseño la casa Fallingwater construida sobre una cascada?', 'nl': 'Welke architect ontwierp het huis Fallingwater, gebouwd over een waterval?', 'sv': 'Vilken arkitekt designade huset Fallingwater byggt over ett vattenfall?', 'fr': "Quel architecte a conçu la maison Fallingwater construite au-dessus d'une cascade ?", 'da': 'Hvilken arkitekt designede huset Fallingwater bygget over et vandfald?'},
 'opt1': {'en': 'Mies van der Rohe', 'de': 'Mies van der Rohe', 'es': 'Mies van der Rohe', 'nl': 'Mies van der Rohe', 'sv': 'Mies van der Rohe', 'fr': 'Mies van der Rohe', 'da': 'Mies van der Rohe'},
 'opt2': {'en': 'Frank Lloyd Wright', 'de': 'Frank Lloyd Wright', 'es': 'Frank Lloyd Wright', 'nl': 'Frank Lloyd Wright', 'sv': 'Frank Lloyd Wright', 'fr': 'Frank Lloyd Wright', 'da': 'Frank Lloyd Wright'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'In Hindu mythology, which goddess is associated with wealth and prosperity?', 'de': 'In der hinduistischen Mythologie, welche Goettin wird mit Reichtum und Wohlstand assoziiert?', 'es': '¿En la mitologia hindu, que diosa se asocia con la riqueza y la prosperidad?', 'nl': 'In de hindoemythologie, welke godin wordt geassocieerd met rijkdom en voorspoed?', 'sv': 'I hinduisk mytologi, vilken gudinna forknippas med rikedom och valstand?', 'fr': "Dans la mythologie hindoue, quelle deesse est associee a la richesse et la prosperite ?", 'da': 'I hinduistisk mytologi, hvilken gudinde forbindes med rigdom og velstand?'},
 'opt1': {'en': 'Saraswati', 'de': 'Saraswati', 'es': 'Saraswati', 'nl': 'Saraswati', 'sv': 'Saraswati', 'fr': 'Saraswati', 'da': 'Saraswati'},
 'opt2': {'en': 'Lakshmi', 'de': 'Lakshmi', 'es': 'Lakshmi', 'nl': 'Lakshmi', 'sv': 'Lakshmi', 'fr': 'Lakshmi', 'da': 'Lakshmi'}},
]
