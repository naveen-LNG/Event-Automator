# -*- coding: utf-8 -*-
# Trivia Challenge batch #8. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which brown sugary syrup is made by boiling down maple tree sap?', 'de': 'Welcher braune zuckrige Sirup wird durch Einkochen von Ahornbaumsaft hergestellt?', 'es': '¿Que jarabe marron y azucarado se hace hirviendo savia de arce?', 'nl': 'Welke bruine suikerige siroop wordt gemaakt door esdoornsap in te koken?', 'sv': 'Vilken brun sockrig sirap gors genom att koka in lonnsav?', 'fr': "Quel sirop brun sucre est fait en faisant bouillir la seve d'erable ?", 'da': 'Hvilken brun sukret sirup er lavet ved at koge ahornsaft ind?'},
 'opt1': {'en': 'Corn syrup', 'de': 'Maissirup', 'es': 'Jarabe de maiz', 'nl': 'Maissiroop', 'sv': 'Majssirap', 'fr': 'Le sirop de mais', 'da': 'Majssirup'},
 'opt2': {'en': 'Maple syrup', 'de': 'Ahornsirup', 'es': 'Jarabe de arce', 'nl': 'Ahornsiroop', 'sv': 'Lonnsirap', 'fr': "Le sirop d'erable", 'da': 'Ahornsirup'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the Great Sphinx?', 'de': 'In welchem Land befindet sich die Grosse Sphinx?', 'es': '¿En que pais esta la Gran Esfinge?', 'nl': 'In welk land bevindt zich de Grote Sfinx?', 'sv': 'I vilket land ligger Sfinxen i Giza?', 'fr': "Dans quel pays se trouve le grand Sphinx ?", 'da': 'Hvilket land huser Den Store Sfinks?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Egypt', 'de': 'Ägypten', 'es': 'Egipto', 'nl': 'Egypte', 'sv': 'Egypten', 'fr': "l'Egypte", 'da': 'Egypten'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic film features a genie who grants three wishes?', 'de': 'In welchem Filmklassiker gewaehrt ein Dschinn drei Wuensche?', 'es': '¿Que pelicula clasica presenta a un genio que concede tres deseos?', 'nl': 'In welke klassieke film verleent een geest drie wensen?', 'sv': 'I vilken klassisk film beviljar en anda tre onskningar?', 'fr': "Quel film classique met en scene un genie qui exauce trois voeux ?", 'da': 'Hvilken klassisk film har en ånd, der opfylder tre ønsker?'},
 'opt1': {'en': 'The Thief of Baghdad', 'de': 'Der Dieb von Bagdad', 'es': 'El ladron de Bagdad', 'nl': 'De Dief van Bagdad', 'sv': 'Tjuven i Bagdad', 'fr': 'Le Voleur de Bagdad', 'da': 'Tyven fra Bagdad'},
 'opt2': {'en': 'Aladdin', 'de': 'Aladdin', 'es': 'Aladdin', 'nl': 'Aladdin', 'sv': 'Aladdin', 'fr': 'Aladdin', 'da': 'Aladdin'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which pet reptile is often kept in a terrarium and can change color?', 'de': 'Welches Haustierreptil wird oft in einem Terrarium gehalten und kann seine Farbe aendern?', 'es': '¿Que reptil mascota se mantiene a menudo en un terrario y puede cambiar de color?', 'nl': 'Welk huisdierreptiel wordt vaak in een terrarium gehouden en kan van kleur veranderen?', 'sv': 'Vilket sallskapsreptil halls ofta i ett terrarium och kan byta farg?', 'fr': "Quel reptile de compagnie est souvent garde dans un terrarium et peut changer de couleur ?", 'da': 'Hvilket kæledyrskrybdyr holdes ofte i et terrarium og kan skifte farve?'},
 'opt1': {'en': 'Chameleon', 'de': 'Chamaeleon', 'es': 'Camaleon', 'nl': 'Kameleon', 'sv': 'Kameleont', 'fr': 'Cameleon', 'da': 'Kamæleon'},
 'opt2': {'en': 'Gecko', 'de': 'Gecko', 'es': 'Gecko', 'nl': 'Gekko', 'sv': 'Gecko', 'fr': 'Gecko', 'da': 'Gekko'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which white cheese is crumbly and commonly used in Greek salads?', 'de': 'Welcher weisse Kaese ist broeselig und wird oft in griechischen Salaten verwendet?', 'es': '¿Que queso blanco es desmenuzable y se usa comunmente en ensaladas griegas?', 'nl': 'Welke witte kaas is kruimelig en wordt vaak gebruikt in Griekse salades?', 'sv': 'Vilken vit ost ar smulig och anvands ofta i grekiska sallader?', 'fr': "Quel fromage blanc est friable et couramment utilise dans les salades grecques ?", 'da': 'Hvilken hvid ost er smuldrende og bruges ofte i græske salater?'},
 'opt1': {'en': 'Ricotta', 'de': 'Ricotta', 'es': 'Ricotta', 'nl': 'Ricotta', 'sv': 'Ricotta', 'fr': 'Ricotta', 'da': 'Ricotta'},
 'opt2': {'en': 'Feta', 'de': 'Feta', 'es': 'Feta', 'nl': 'Feta', 'sv': 'Feta', 'fr': 'Feta', 'da': 'Feta'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to Niagara Falls, shared with Canada?', 'de': 'In welchem Land liegen die Niagarafaelle, geteilt mit Kanada?', 'es': '¿En que pais estan las Cataratas del Niagara, compartidas con Canada?', 'nl': 'In welk land liggen Niagara Falls, gedeeld met Canada?', 'sv': 'I vilket land ligger Niagarafallen, delade med Kanada?', 'fr': "Dans quel pays se trouvent les chutes du Niagara, partagees avec le Canada ?", 'da': 'Hvilket land huser Niagara Falls, delt med Canada?'},
 'opt1': {'en': 'United States', 'de': 'Vereinigte Staaten', 'es': 'Estados Unidos', 'nl': 'Verenigde Staten', 'sv': 'USA', 'fr': 'les États-Unis', 'da': 'USA'},
 'opt2': {'en': 'Mexico', 'de': 'Mexiko', 'es': 'Mexico', 'nl': 'Mexico', 'sv': 'Mexiko', 'fr': 'le Mexique', 'da': 'Mexico'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. president served during World War II and died before it ended?', 'de': 'Welcher US-Praesident diente waehrend des Zweiten Weltkriegs und starb, bevor er endete?', 'es': '¿Que presidente de EE.UU. sirvio durante la Segunda Guerra Mundial y murio antes de que terminara?', 'nl': 'Welke Amerikaanse president diende tijdens de Tweede Wereldoorlog en stierf voordat deze eindigde?', 'sv': 'Vilken amerikansk president tjanstgjorde under andra varldskriget och dog innan det tog slut?', 'fr': "Quel president americain a servi pendant la Seconde Guerre mondiale et est mort avant sa fin ?", 'da': 'Hvilken amerikansk præsident tjente under 2. verdenskrig og døde, før den sluttede?'},
 'opt1': {'en': 'Harry Truman', 'de': 'Harry Truman', 'es': 'Harry Truman', 'nl': 'Harry Truman', 'sv': 'Harry Truman', 'fr': 'Harry Truman', 'da': 'Harry Truman'},
 'opt2': {'en': 'Franklin D. Roosevelt', 'de': 'Franklin D. Roosevelt', 'es': 'Franklin D. Roosevelt', 'nl': 'Franklin D. Roosevelt', 'sv': 'Franklin D. Roosevelt', 'fr': 'Franklin D. Roosevelt', 'da': 'Franklin D. Roosevelt'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which craft involves shaping clay into pots using a spinning wheel?', 'de': 'Bei welchem Handwerk wird Ton mit einer Drehscheibe zu Toepfen geformt?', 'es': '¿Que manualidad implica dar forma a la arcilla en ollas usando un torno giratorio?', 'nl': 'Bij welk handwerk vorm je klei tot potten met een draaischijf?', 'sv': 'Vilket hantverk innebar att forma lera till krukor med en snurrande skiva?', 'fr': "Quel art manuel consiste a façonner de l'argile en pots avec un tour ?", 'da': 'Hvilket håndværk involverer at forme ler til potter ved hjælp af en drejeskive?'},
 'opt1': {'en': 'Ceramics', 'de': 'Keramik', 'es': 'Ceramica', 'nl': 'Keramiek', 'sv': 'Keramikkonst', 'fr': 'La ceramique', 'da': 'Keramikkunst'},
 'opt2': {'en': 'Pottery', 'de': 'Toepferei', 'es': 'Alfareria', 'nl': 'Pottenbakken', 'sv': 'Keramik', 'fr': 'La poterie', 'da': 'Keramik'}},

{'category': 'Music', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which stringed instrument is plucked and often used in classical Spanish music?', 'de': 'Welches Saiteninstrument wird gezupft und oft in klassischer spanischer Musik verwendet?', 'es': '¿Que instrumento de cuerda se puntea y se usa a menudo en musica clasica española?', 'nl': 'Welk snaarinstrument wordt getokkeld en vaak gebruikt in klassieke Spaanse muziek?', 'sv': 'Vilket strakinstrument knapps och anvands ofta i klassisk spansk musik?', 'fr': "Quel instrument a cordes se pince et est souvent utilise dans la musique classique espagnole ?", 'da': 'Hvilket strengeinstrument klimpres og bruges ofte i klassisk spansk musik?'},
 'opt1': {'en': 'Lute', 'de': 'Laute', 'es': 'Laud', 'nl': 'Luit', 'sv': 'Luta', 'fr': 'Luth', 'da': 'Lut'},
 'opt2': {'en': 'Guitar', 'de': 'Gitarre', 'es': 'Guitarra', 'nl': 'Gitaar', 'sv': 'Gitarr', 'fr': 'Guitare', 'da': 'Guitar'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which weather phenomenon produces thunder and lightning?', 'de': 'Welches Wetterphaenomen erzeugt Donner und Blitz?', 'es': '¿Que fenomeno meteorologico produce truenos y relampagos?', 'nl': 'Welk weersverschijnsel produceert donder en bliksem?', 'sv': 'Vilket vaderfenomen producerar aska och blixt?', 'fr': "Quel phenomene meteorologique produit du tonnerre et des eclairs ?", 'da': 'Hvilket vejrfænomen producerer torden og lyn?'},
 'opt1': {'en': 'Thunderstorm', 'de': 'Gewitter', 'es': 'Tormenta electrica', 'nl': 'Onweer', 'sv': 'Askvader', 'fr': 'Orage', 'da': 'Tordenvejr'},
 'opt2': {'en': 'Hailstorm', 'de': 'Hagelsturm', 'es': 'Granizada', 'nl': 'Hagelbui', 'sv': 'Hagelstorm', 'fr': 'Tempete de grele', 'da': 'Haglstorm'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a robot who falls in love while cleaning up Earth?', 'de': 'In welchem Filmklassiker verliebt sich ein Roboter, waehrend er die Erde aufraeumt?', 'es': '¿Que pelicula clasica presenta a un robot que se enamora mientras limpia la Tierra?', 'nl': 'In welke klassieke film wordt een robot verliefd terwijl hij de aarde opruimt?', 'sv': 'I vilken klassisk film blir en robot forälskad medan den stadar upp jorden?', 'fr': "Quel film classique met en scene un robot qui tombe amoureux en nettoyant la Terre ?", 'da': 'Hvilken klassisk film har en robot, der bliver forelsket, mens den rydder op på Jorden?'},
 'opt1': {'en': 'WALL-E', 'de': 'WALL-E', 'es': 'WALL-E', 'nl': 'WALL-E', 'sv': 'WALL-E', 'fr': 'WALL-E', 'da': 'WALL-E'},
 'opt2': {'en': 'Short Circuit', 'de': 'Nummer 5 lebt', 'es': 'Cortocircuito', 'nl': 'Short Circuit', 'sv': 'Short Circuit', 'fr': 'Numero 5 ne se recharge pas', 'da': 'Short Circuit'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which U.S. state is known as the 'Pelican State'?", 'de': "Welcher US-Bundesstaat ist als 'Pelican State' bekannt?", 'es': "¿Que estado de EE.UU. es conocido como el 'Estado del Pelicano'?", 'nl': "Welke Amerikaanse staat staat bekend als de 'Pelican State'?", 'sv': "Vilken amerikansk delstat ar kand som 'Pelikanstaten'?", 'fr': "Quel État americain est connu comme le 'Pelican State' (État du pelican) ?", 'da': "Hvilken amerikansk delstat er kendt som 'Pelican State'?"},
 'opt1': {'en': 'Louisiana', 'de': 'Louisiana', 'es': 'Luisiana', 'nl': 'Louisiana', 'sv': 'Louisiana', 'fr': 'la Louisiane', 'da': 'Louisiana'},
 'opt2': {'en': 'Mississippi', 'de': 'Mississippi', 'es': 'Misisipi', 'nl': 'Mississippi', 'sv': 'Mississippi', 'fr': 'le Mississippi', 'da': 'Mississippi'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which green fruit is used to make guacamole?', 'de': 'Welche gruene Frucht wird fuer Guacamole verwendet?', 'es': '¿Que fruta verde se usa para hacer guacamole?', 'nl': 'Welk groen fruit wordt gebruikt om guacamole te maken?', 'sv': 'Vilken gron frukt anvands for att gora guacamole?', 'fr': "Quel fruit vert est utilise pour faire du guacamole ?", 'da': 'Hvilken grøn frugt bruges til at lave guacamole?'},
 'opt1': {'en': 'Kiwi', 'de': 'Kiwi', 'es': 'Kiwi', 'nl': 'Kiwi', 'sv': 'Kiwi', 'fr': 'Le kiwi', 'da': 'Kiwi'},
 'opt2': {'en': 'Avocado', 'de': 'Avocado', 'es': 'Aguacate', 'nl': 'Avocado', 'sv': 'Avokado', 'fr': "L'avocat", 'da': 'Avocado'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves hitting a shuttlecock over a net with a racquet?', 'de': 'Bei welcher Sportart schlaegt man einen Federball mit einem Schlaeger ueber ein Netz?', 'es': '¿Que deporte consiste en golpear un volante sobre una red con una raqueta?', 'nl': 'Bij welke sport sla je een shuttle met een racket over een net?', 'sv': 'Vilken sport innebar att slå en badmintonboll over ett nat med en racket?', 'fr': "Quel sport consiste a frapper un volant par-dessus un filet avec une raquette ?", 'da': 'Hvilken sport involverer at slå en fjerbold over et net med en ketcher?'},
 'opt1': {'en': 'Badminton', 'de': 'Badminton', 'es': 'Badminton', 'nl': 'Badminton', 'sv': 'Badminton', 'fr': 'Le badminton', 'da': 'Badminton'},
 'opt2': {'en': 'Tennis', 'de': 'Tennis', 'es': 'Tenis', 'nl': 'Tennis', 'sv': 'Tennis', 'fr': 'Tennis', 'da': 'Tennis'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which pet fish is commonly kept in a bowl and comes in orange varieties?', 'de': 'Welcher Haustierfisch wird oft in einer Schuessel gehalten und kommt in orangefarbenen Varianten vor?', 'es': '¿Que pez mascota se mantiene comunmente en un tazon y viene en variedades naranjas?', 'nl': 'Welke huisdiervis wordt vaak in een kom gehouden en komt in oranje varianten voor?', 'sv': 'Vilken sallskapsfisk halls vanligtvis i en skal och finns i oranga varianter?', 'fr': "Quel poisson de compagnie est couramment garde dans un bocal et existe en variantes orange ?", 'da': 'Hvilken kæledyrsfisk holdes almindeligvis i en skål og fås i orange varianter?'},
 'opt1': {'en': 'Betta fish', 'de': 'Kampffisch', 'es': 'Pez betta', 'nl': 'Kempvis', 'sv': 'Kampfisk', 'fr': 'Poisson combattant', 'da': 'Kampfisk'},
 'opt2': {'en': 'Goldfish', 'de': 'Goldfisch', 'es': 'Pez dorado', 'nl': 'Goudvis', 'sv': 'Guldfisk', 'fr': 'Poisson rouge', 'da': 'Guldfisk'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which holiday celebrates love with cards and chocolates in February?', 'de': 'Welcher Feiertag im Februar feiert die Liebe mit Karten und Schokolade?', 'es': '¿Que festividad de febrero celebra el amor con tarjetas y chocolates?', 'nl': 'Welke feestdag in februari viert de liefde met kaarten en chocolade?', 'sv': 'Vilken hogtid i februari firar karlek med kort och choklad?', 'fr': "Quelle fete de fevrier celebre l'amour avec des cartes et du chocolat ?", 'da': 'Hvilken helligdag i februar fejrer kærlighed med kort og chokolade?'},
 'opt1': {'en': "Sweetest Day", 'de': 'Sweetest Day', 'es': 'Sweetest Day', 'nl': 'Sweetest Day', 'sv': 'Sweetest Day', 'fr': 'Sweetest Day', 'da': 'Sweetest Day'},
 'opt2': {'en': "Valentine's Day", 'de': 'Valentinstag', 'es': 'Dia de San Valentin', 'nl': 'Valentijnsdag', 'sv': 'Alla hjartans dag', 'fr': 'La Saint-Valentin', 'da': 'Valentinsdag'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which continent is home to Mount Kilimanjaro?', 'de': 'Auf welchem Kontinent liegt der Kilimandscharo?', 'es': '¿En que continente esta el Monte Kilimanjaro?', 'nl': 'Op welk continent ligt de Kilimanjaro?', 'sv': 'Pa vilken kontinent ligger Kilimanjaro?', 'fr': "Sur quel continent se trouve le mont Kilimandjaro ?", 'da': 'Hvilket kontinent huser Kilimanjaro-bjerget?'},
 'opt1': {'en': 'Asia', 'de': 'Asien', 'es': 'Asia', 'nl': 'Azie', 'sv': 'Asien', 'fr': "l'Asie", 'da': 'Asien'},
 'opt2': {'en': 'Africa', 'de': 'Afrika', 'es': 'Africa', 'nl': 'Afrika', 'sv': 'Afrika', 'fr': "l'Afrique", 'da': 'Afrika'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which author wrote 'Matilda'?", 'de': "Welcher Autor schrieb 'Matilda'?", 'es': "¿Que autor escribio 'Matilda'?", 'nl': "Welke auteur schreef 'Matilda'?", 'sv': "Vilken forfattare skrev 'Matilda'?", 'fr': "Quel auteur a ecrit 'Matilda' ?", 'da': "Hvilken forfatter skrev 'Matilda'?"},
 'opt1': {'en': 'Roald Dahl', 'de': 'Roald Dahl', 'es': 'Roald Dahl', 'nl': 'Roald Dahl', 'sv': 'Roald Dahl', 'fr': 'Roald Dahl', 'da': 'Roald Dahl'},
 'opt2': {'en': 'Dr. Seuss', 'de': 'Dr. Seuss', 'es': 'Dr. Seuss', 'nl': 'Dr. Seuss', 'sv': 'Dr. Seuss', 'fr': 'Dr Seuss', 'da': 'Dr. Seuss'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic toy involves stacking rings of decreasing size on a pole?', 'de': 'Bei welchem klassischen Spielzeug stapelt man Ringe abnehmender Groesse auf einen Pfahl?', 'es': '¿Que juguete clasico consiste en apilar anillos de tamaño decreciente en un poste?', 'nl': 'Bij welk klassiek speelgoed stapel je ringen van afnemende grootte op een paal?', 'sv': 'Vilken klassisk leksak innebar att stapla ringar av minskande storlek pa en stang?', 'fr': "Quel jouet classique consiste a empiler des anneaux de taille decroissante sur un poteau ?", 'da': 'Hvilket klassisk legetøj involverer at stable ringe af aftagende størrelse på en stang?'},
 'opt1': {'en': 'Stacking rings toy', 'de': 'Stapelringspielzeug', 'es': 'Aros apilables', 'nl': 'Stapelringen', 'sv': 'Stapelringar', 'fr': "Anneaux empilables", 'da': 'Stableringe'},
 'opt2': {'en': 'Building blocks', 'de': 'Bauklötze', 'es': 'Bloques de construccion', 'nl': 'Bouwblokken', 'sv': 'Byggklossar', 'fr': 'Blocs de construction', 'da': 'Byggeklodser'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which white spice is ground from a tropical seed and used in eggnog?', 'de': 'Welches weisse Gewuerz wird aus einem tropischen Samen gemahlen und in Eggnog verwendet?', 'es': '¿Que especia blanca se muele de una semilla tropical y se usa en el rompope?', 'nl': 'Welke witte specerij wordt gemalen van een tropisch zaad en gebruikt in eggnog?', 'sv': 'Vilken vit krydda mals fran ett tropiskt fro och anvands i aggtoddy?', 'fr': "Quelle epice blanche est moulue a partir d'une graine tropicale et utilisee dans le lait de poule ?", 'da': 'Hvilket hvidt krydderi males af et tropisk frø og bruges i æggesnaps?'},
 'opt1': {'en': 'Cinnamon', 'de': 'Zimt', 'es': 'Canela', 'nl': 'Kaneel', 'sv': 'Kanel', 'fr': 'La cannelle', 'da': 'Kanel'},
 'opt2': {'en': 'Nutmeg', 'de': 'Muskatnuss', 'es': 'Nuez moscada', 'nl': 'Nootmuskaat', 'sv': 'Muskot', 'fr': 'La muscade', 'da': 'Muskatnød'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. holiday celebrates the signing of the Declaration of Independence?', 'de': 'Welcher US-Feiertag feiert die Unterzeichnung der Unabhaengigkeitserklaerung?', 'es': '¿Que festividad de EE.UU. celebra la firma de la Declaracion de Independencia?', 'nl': 'Welke Amerikaanse feestdag viert de ondertekening van de Onafhankelijkheidsverklaring?', 'sv': 'Vilken amerikansk hogtid firar undertecknandet av sjalvstandighetsforklaringen?', 'fr': "Quelle fete americaine celebre la signature de la Declaration d'independance ?", 'da': 'Hvilken amerikansk helligdag fejrer underskrivelsen af uafhængighedserklæringen?'},
 'opt1': {'en': 'Constitution Day', 'de': 'Verfassungstag', 'es': 'Dia de la Constitucion', 'nl': 'Grondwetdag', 'sv': 'Konstitutionsdagen', 'fr': 'Le jour de la Constitution', 'da': 'Forfatningsdag'},
 'opt2': {'en': 'Independence Day', 'de': 'Independence Day', 'es': 'Dia de la Independencia', 'nl': 'Independence Day', 'sv': 'Sjalvstandighetsdagen', 'fr': "Le jour de l'Independance", 'da': 'Uafhængighedsdag'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which decade is associated with swing and big band music?', 'de': 'Mit welchem Jahrzehnt wird Swing- und Big-Band-Musik assoziiert?', 'es': '¿Con que decada se asocia la musica swing y de big band?', 'nl': 'Met welk decennium wordt swing- en bigbandmuziek geassocieerd?', 'sv': 'Vilket artionde forknippas med swing- och storbandsmusik?', 'fr': "A quelle decennie associe-t-on la musique swing et big band ?", 'da': 'Hvilket årti forbindes med swing- og big band-musik?'},
 'opt1': {'en': '1940s', 'de': '1940er', 'es': 'los años 40', 'nl': 'jaren 40', 'sv': '1940-talet', 'fr': 'les annees 1940', 'da': '1940’erne'},
 'opt2': {'en': '1920s', 'de': '1920er', 'es': 'los años 20', 'nl': 'jaren 20', 'sv': '1920-talet', 'fr': 'les annees 1920', 'da': '1920’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which country is home to the ancient city of Bogota's neighboring salt cathedral of Zipaquira?", 'de': 'In welchem Land liegt die Salzkathedrale von Zipaquira, nahe Bogota?', 'es': '¿En que pais esta la catedral de sal de Zipaquira, cerca de Bogota?', 'nl': 'In welk land ligt de zoutkathedraal van Zipaquira, nabij Bogota?', 'sv': 'I vilket land ligger saltkatedralen i Zipaquira, nara Bogota?', 'fr': "Dans quel pays se trouve la cathedrale de sel de Zipaquira, pres de Bogota ?", 'da': 'Hvilket land huser saltkatedralen i Zipaquira, nær Bogotá?'},
 'opt1': {'en': 'Ecuador', 'de': 'Ecuador', 'es': 'Ecuador', 'nl': 'Ecuador', 'sv': 'Ecuador', 'fr': "l'Equateur", 'da': 'Ecuador'},
 'opt2': {'en': 'Colombia', 'de': 'Kolumbien', 'es': 'Colombia', 'nl': 'Colombia', 'sv': 'Colombia', 'fr': 'la Colombie', 'da': 'Colombia'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which farm animal is known for clucking and laying eggs?', 'de': 'Welches Nutztier gackert und legt Eier?', 'es': '¿Que animal de granja cacarea y pone huevos?', 'nl': 'Welk boerderijdier kakelt en legt eieren?', 'sv': 'Vilket lantbruksdjur kacklar och lagger agg?', 'fr': "Quel animal de ferme caquette et pond des œufs ?", 'da': 'Hvilket landbrugsdyr klukker og lægger æg?'},
 'opt1': {'en': 'Chicken', 'de': 'Huhn', 'es': 'Gallina', 'nl': 'Kip', 'sv': 'Höna', 'fr': 'Poule', 'da': 'Høne'},
 'opt2': {'en': 'Duck', 'de': 'Ente', 'es': 'Pato', 'nl': 'Eend', 'sv': 'And', 'fr': 'Canard', 'da': 'And'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which gemstone is the birthstone for those born in October?', 'de': 'Welcher Edelstein ist der Geburtsstein fuer im Oktober Geborene?', 'es': '¿Que piedra preciosa es la piedra de nacimiento para los nacidos en octubre?', 'nl': 'Welke edelsteen is de geboortesteen voor mensen geboren in oktober?', 'sv': 'Vilken adelsten ar fodelsestenen for de fodda i oktober?', 'fr': "Quelle pierre precieuse est la pierre de naissance d'octobre ?", 'da': 'Hvilken ædelsten er fødselsstenen for dem, der er født i oktober?'},
 'opt1': {'en': 'Pearl', 'de': 'Perle', 'es': 'Perla', 'nl': 'Parel', 'sv': 'Parla', 'fr': 'Perle', 'da': 'Perle'},
 'opt2': {'en': 'Opal', 'de': 'Opal', 'es': 'Opalo', 'nl': 'Opaal', 'sv': 'Opal', 'fr': 'Opale', 'da': 'Opal'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic outdoor game involves throwing horseshoes or rings at a stake?', 'de': 'Bei welchem klassischen Freiluftspiel wirft man Hufeisen oder Ringe auf einen Pfahl?', 'es': '¿Que juego clasico al aire libre implica lanzar herraduras o aros a una estaca?', 'nl': 'Bij welk klassiek buitenspel gooi je hoefijzers of ringen naar een paal?', 'sv': 'Vilket klassiskt utomhusspel innebar att kasta hastskor eller ringar mot en pale?', 'fr': "Quel jeu de plein air classique consiste a lancer des fers a cheval ou des anneaux sur un piquet ?", 'da': 'Hvilket klassisk udendørsspil involverer at kaste hesteskoe eller ringe mod en pæl?'},
 'opt1': {'en': 'Ring toss', 'de': 'Ringwerfen', 'es': 'Lanzamiento de aros', 'nl': 'Ringwerpen', 'sv': 'Ringkastning', 'fr': "Le lancer d'anneaux", 'da': 'Ringkast'},
 'opt2': {'en': 'Cornhole', 'de': 'Cornhole', 'es': 'Cornhole', 'nl': 'Cornhole', 'sv': 'Cornhole', 'fr': 'Cornhole', 'da': 'Cornhole'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which green herb is used to flavor traditional Italian tomato sauce?', 'de': 'Welches gruene Kraut wird verwendet, um traditionelle italienische Tomatensauce zu wuerzen?', 'es': '¿Que hierba verde se usa para sazonar la salsa de tomate italiana tradicional?', 'nl': 'Welk groen kruid wordt gebruikt om traditionele Italiaanse tomatensaus te kruiden?', 'sv': 'Vilken gron ort anvands for att smaksatta traditionell italiensk tomatsas?', 'fr': "Quelle herbe verte est utilisee pour aromatiser la sauce tomate italienne traditionnelle ?", 'da': 'Hvilken grøn urt bruges til at smage traditionel italiensk tomatsauce?'},
 'opt1': {'en': 'Oregano', 'de': 'Oregano', 'es': 'Oregano', 'nl': 'Oregano', 'sv': 'Oregano', 'fr': "L'origan", 'da': 'Oregano'},
 'opt2': {'en': 'Thyme', 'de': 'Thymian', 'es': 'Tomillo', 'nl': 'Tijm', 'sv': 'Timjan', 'fr': 'Le thym', 'da': 'Timian'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which flower is associated with the Chinese New Year for good fortune?', 'de': 'Welche Blume wird mit dem chinesischen Neujahr fuer Glueck assoziiert?', 'es': '¿Que flor se asocia con el Año Nuevo Chino para la buena fortuna?', 'nl': 'Welke bloem wordt geassocieerd met het Chinees Nieuwjaar voor geluk?', 'sv': 'Vilken blomma forknippas med kinesiskt nyar for god lycka?', 'fr': "Quelle fleur est associee au Nouvel An chinois pour la bonne fortune ?", 'da': 'Hvilken blomst forbindes med det kinesiske nytår for held?'},
 'opt1': {'en': 'Peony', 'de': 'Pfingstrose', 'es': 'Peonia', 'nl': 'Pioenroos', 'sv': 'Pion', 'fr': 'La pivoine', 'da': 'Pæon'},
 'opt2': {'en': 'Chrysanthemum', 'de': 'Chrysantheme', 'es': 'Crisantemo', 'nl': 'Chrysant', 'sv': 'Krysantemum', 'fr': 'Chrysantheme', 'da': 'Krysantemum'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Liberty Bell in Philadelphia?', 'de': 'In welchem US-Bundesstaat liegt die Freiheitsglocke in Philadelphia?', 'es': '¿En que estado de EE.UU. esta la Campana de la Libertad en Filadelfia?', 'nl': 'In welke Amerikaanse staat ligt de Liberty Bell in Philadelphia?', 'sv': 'I vilken amerikansk delstat ligger Frihetsklockan i Philadelphia?', 'fr': "Dans quel État americain se trouve la cloche de la Liberte a Philadelphie ?", 'da': 'I hvilken amerikansk delstat ligger Frihedsklokken i Philadelphia?'},
 'opt1': {'en': 'Pennsylvania', 'de': 'Pennsylvania', 'es': 'Pensilvania', 'nl': 'Pennsylvania', 'sv': 'Pennsylvania', 'fr': 'la Pennsylvanie', 'da': 'Pennsylvania'},
 'opt2': {'en': 'New Jersey', 'de': 'New Jersey', 'es': 'Nueva Jersey', 'nl': 'New Jersey', 'sv': 'New Jersey', 'fr': 'le New Jersey', 'da': 'New Jersey'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which craft involves stitching designs onto fabric using colored thread?', 'de': 'Bei welchem Handwerk werden Muster mit farbigem Faden auf Stoff gestickt?', 'es': '¿Que manualidad implica coser diseños en tela usando hilo de color?', 'nl': 'Bij welk handwerk stik je patronen op stof met gekleurde draad?', 'sv': 'Vilket hantverk innebar att sy monster pa tyg med fargad trad?', 'fr': "Quel art manuel consiste a broder des motifs sur du tissu avec du fil colore ?", 'da': 'Hvilket håndværk involverer at sy mønstre på stof med farvet tråd?'},
 'opt1': {'en': 'Quilting', 'de': 'Quilten', 'es': 'Acolchado', 'nl': 'Quilten', 'sv': 'Lapptacksarbete', 'fr': 'Le patchwork', 'da': 'Quiltning'},
 'opt2': {'en': 'Embroidery', 'de': 'Stickerei', 'es': 'Bordado', 'nl': 'Borduurwerk', 'sv': 'Broderi', 'fr': 'Broderie', 'da': 'Broderi'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which sport involves hitting a ball across a net using only the head, feet, and chest?', 'de': 'Bei welcher Sportart wird ein Ball nur mit Kopf, Fuessen und Brust ueber ein Netz gespielt?', 'es': '¿Que deporte consiste en golpear una pelota sobre una red usando solo la cabeza, los pies y el pecho?', 'nl': 'Bij welke sport speel je een bal over een net met alleen hoofd, voeten en borst?', 'sv': 'Vilken sport innebar att slå en boll over ett nat med bara huvud, fotter och brost?', 'fr': "Quel sport consiste a envoyer un ballon par-dessus un filet en utilisant seulement la tete, les pieds et la poitrine ?", 'da': 'Hvilken sport involverer at sende en bold over et net kun med hovedet, fødderne og brystet?'},
 'opt1': {'en': 'Volleyball', 'de': 'Volleyball', 'es': 'Voleibol', 'nl': 'Volleybal', 'sv': 'Volleyboll', 'fr': 'Le volley-ball', 'da': 'Volleyball'},
 'opt2': {'en': 'Soccer (football)', 'de': 'Fussball', 'es': 'Futbol', 'nl': 'Voetbal', 'sv': 'Fotboll', 'fr': 'Le football', 'da': 'Fodbold'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which item of clothing is worn over the shoulders to keep warm, often knitted?', 'de': 'Welches Kleidungsstueck wird zum Waermen ueber die Schultern getragen, oft gestrickt?', 'es': '¿Que prenda se lleva sobre los hombros para mantener el calor, a menudo tejida?', 'nl': 'Welk kledingstuk wordt over de schouders gedragen om warm te blijven, vaak gebreid?', 'sv': 'Vilket klädesplagg baers over axlarna for att halla varmen, ofta stickat?', 'fr': "Quel vetement se porte sur les epaules pour se tenir chaud, souvent tricote ?", 'da': 'Hvilken beklædningsgenstand bæres over skuldrene for at holde varmen, ofte strikket?'},
 'opt1': {'en': 'Shawl', 'de': 'Schal', 'es': 'Chal', 'nl': 'Omslagdoek', 'sv': 'Sjal', 'fr': 'Chale', 'da': 'Sjal'},
 'opt2': {'en': 'Scarf', 'de': 'Schal (Halstuch)', 'es': 'Bufanda', 'nl': 'Sjaal', 'sv': 'Halsduk', 'fr': 'Echarpe', 'da': 'Tørklæde'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic cartoon features a duck who wears a sailor suit and has a bad temper?', 'de': 'In welchem klassischen Cartoon kommt eine Ente vor, die einen Matrosenanzug traegt und schnell wuetend wird?', 'es': '¿Que dibujo animado clasico presenta a un pato que lleva traje de marinero y tiene mal genio?', 'nl': 'In welke klassieke tekenfilm draagt een eend een matrozenpak en heeft hij een kort lontje?', 'sv': 'I vilken klassisk tecknad film forekommer en anka som baer en sjomansdrakt och har ett kort humor?', 'fr': "Quel dessin anime classique met en scene un canard qui porte un costume de marin et a mauvais caractere ?", 'da': 'Hvilken klassisk tegnefilm har en and, der bærer et matrosdragt og har hidsigt temperament?'},
 'opt1': {'en': 'Donald Duck', 'de': 'Donald Duck', 'es': 'Pato Donald', 'nl': 'Donald Duck', 'sv': 'Kalle Anka', 'fr': 'Donald Duck', 'da': 'Anders And'},
 'opt2': {'en': 'Daffy Duck', 'de': 'Daffy Duck', 'es': 'Pato Lucas', 'nl': 'Daffy Duck', 'sv': 'Daffy Anka', 'fr': 'Daffy Duck', 'da': 'Daffy And'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Petra, carved into rock?', 'de': 'In welchem Land liegt die antike, in Fels gehauene Stadt Petra?', 'es': '¿En que pais esta la antigua ciudad de Petra, tallada en roca?', 'nl': 'In welk land ligt de oude, uit rots gehouwen stad Petra?', 'sv': 'I vilket land ligger den antika klipphuggna staden Petra?', 'fr': "Dans quel pays se trouve l'ancienne cite de Petra, taillee dans la roche ?", 'da': 'Hvilket land huser den antikke klippehuggede by Petra?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Egypt', 'de': 'Ägypten', 'es': 'Egipto', 'nl': 'Egypte', 'sv': 'Egypten', 'fr': "l'Egypte", 'da': 'Egypten'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which First Lady founded the Head Start program's early supporters through her War on Poverty advocacy?", 'de': 'Welche First Lady gehoerte durch ihr Engagement im Krieg gegen die Armut zu den fruehen Unterstuetzern des Head-Start-Programms?', 'es': '¿Que primera dama fue una de las primeras impulsoras del programa Head Start a traves de su defensa de la Guerra contra la Pobreza?', 'nl': 'Welke First Lady behoorde tot de vroege ondersteuners van het Head Start-programma via haar pleidooi in de Oorlog tegen Armoede?', 'sv': 'Vilken First Lady var en tidig forespakare for Head Start-programmet genom sitt engagemang mot fattigdom?', 'fr': "Quelle Première dame a ete l'une des premieres a soutenir le programme Head Start a travers son plaidoyer pour la guerre contre la pauvrete ?", 'da': 'Hvilken førstedame var en tidlig fortaler for Head Start-programmet gennem sit arbejde mod fattigdom?'},
 'opt1': {'en': 'Eleanor Roosevelt', 'de': 'Eleanor Roosevelt', 'es': 'Eleanor Roosevelt', 'nl': 'Eleanor Roosevelt', 'sv': 'Eleanor Roosevelt', 'fr': 'Eleanor Roosevelt', 'da': 'Eleanor Roosevelt'},
 'opt2': {'en': 'Lady Bird Johnson', 'de': 'Lady Bird Johnson', 'es': 'Lady Bird Johnson', 'nl': 'Lady Bird Johnson', 'sv': 'Lady Bird Johnson', 'fr': 'Lady Bird Johnson', 'da': 'Lady Bird Johnson'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Lebanese dish is a salad made from bulgur wheat, parsley, and tomatoes?', 'de': 'Welches libanesische Gericht ist ein Salat aus Bulgur, Petersilie und Tomaten?', 'es': '¿Que plato libanes es una ensalada hecha de trigo bulgur, perejil y tomates?', 'nl': 'Welk Libanees gerecht is een salade van bulgurtarwe, peterselie en tomaten?', 'sv': 'Vilken libanesisk ratt ar en sallad gjord av bulgurvete, persilja och tomater?', 'fr': "Quel plat libanais est une salade a base de boulgour, de persil et de tomates ?", 'da': 'Hvilken libanesisk ret er en salat lavet af bulgurhvede, persille og tomater?'},
 'opt1': {'en': 'Tabbouleh', 'de': 'Taboule', 'es': 'Tabule', 'nl': 'Tabouleh', 'sv': 'Tabbouleh', 'fr': 'Taboule', 'da': 'Tabbouleh'},
 'opt2': {'en': 'Fattoush', 'de': 'Fattoush', 'es': 'Fattoush', 'nl': 'Fattoush', 'sv': 'Fattoush', 'fr': 'Fattouche', 'da': 'Fattoush'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which organ is responsible for filtering waste and producing urine?', 'de': 'Welches Organ ist fuer die Filterung von Abfallstoffen und die Urinproduktion verantwortlich?', 'es': '¿Que organo es responsable de filtrar los desechos y producir orina?', 'nl': 'Welk orgaan is verantwoordelijk voor het filteren van afval en het produceren van urine?', 'sv': 'Vilket organ ansvarar for att filtrera avfall och producera urin?', 'fr': "Quel organe est responsable de filtrer les dechets et de produire l'urine ?", 'da': 'Hvilket organ er ansvarlig for at filtrere affald og producere urin?'},
 'opt1': {'en': 'Bladder', 'de': 'Blase', 'es': 'Vejiga', 'nl': 'Blaas', 'sv': 'Blase', 'fr': 'Vessie', 'da': 'Blære'},
 'opt2': {'en': 'Kidney', 'de': 'Niere', 'es': 'Riñon', 'nl': 'Nier', 'sv': 'Njure', 'fr': 'Rein', 'da': 'Nyre'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'The Color Purple'?", 'de': "Welche Autorin schrieb 'Die Farbe Lila'?", 'es': "¿Que autora escribio 'El color purpura'?", 'nl': "Welke auteur schreef 'The Color Purple'?", 'sv': "Vilken forfattare skrev 'Fargen lila'?", 'fr': "Quelle autrice a ecrit 'La Couleur pourpre' ?", 'da': "Hvilken forfatter skrev 'Purpurfarven'?"},
 'opt1': {'en': 'Toni Morrison', 'de': 'Toni Morrison', 'es': 'Toni Morrison', 'nl': 'Toni Morrison', 'sv': 'Toni Morrison', 'fr': 'Toni Morrison', 'da': 'Toni Morrison'},
 'opt2': {'en': 'Alice Walker', 'de': 'Alice Walker', 'es': 'Alice Walker', 'nl': 'Alice Walker', 'sv': 'Alice Walker', 'fr': 'Alice Walker', 'da': 'Alice Walker'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Sarah Sanderson in the film 'Hocus Pocus'?", 'de': "Welche Schauspielerin spielte Sarah Sanderson im Film 'Hokuspokus'?", 'es': "¿Que actriz interpreto a Sarah Sanderson en la pelicula 'Hocus Pocus'?", 'nl': "Welke actrice speelde Sarah Sanderson in de film 'Hocus Pocus'?", 'sv': "Vilken skadespelerska spelade Sarah Sanderson i filmen 'Hocus Pocus'?", 'fr': "Quelle actrice a joue Sarah Sanderson dans le film 'Hocus Pocus' ?", 'da': "Hvilken skuespiller spillede Sarah Sanderson i filmen 'Hocus Pocus'?"},
 'opt1': {'en': 'Kathy Najimy', 'de': 'Kathy Najimy', 'es': 'Kathy Najimy', 'nl': 'Kathy Najimy', 'sv': 'Kathy Najimy', 'fr': 'Kathy Najimy', 'da': 'Kathy Najimy'},
 'opt2': {'en': 'Sarah Jessica Parker', 'de': 'Sarah Jessica Parker', 'es': 'Sarah Jessica Parker', 'nl': 'Sarah Jessica Parker', 'sv': 'Sarah Jessica Parker', 'fr': 'Sarah Jessica Parker', 'da': 'Sarah Jessica Parker'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which singer is known for the song 'I Will Always Love You', originally written by Dolly Parton?", 'de': "Welche Saengerin ist fuer das Lied 'I Will Always Love You' bekannt, urspruenglich geschrieben von Dolly Parton?", 'es': "¿Que cantante es conocida por la cancion 'I Will Always Love You', escrita originalmente por Dolly Parton?", 'nl': "Welke zangeres staat bekend om het lied 'I Will Always Love You', oorspronkelijk geschreven door Dolly Parton?", 'sv': "Vilken sangerska ar kand for latten 'I Will Always Love You', ursprungligen skriven av Dolly Parton?", 'fr': "Quelle chanteuse est connue pour la chanson 'I Will Always Love You', ecrite a l'origine par Dolly Parton ?", 'da': "Hvilken sanger er kendt for sangen 'I Will Always Love You', oprindeligt skrevet af Dolly Parton?"},
 'opt1': {'en': 'Celine Dion', 'de': 'Celine Dion', 'es': 'Celine Dion', 'nl': 'Celine Dion', 'sv': 'Celine Dion', 'fr': 'Celine Dion', 'da': 'Celine Dion'},
 'opt2': {'en': 'Whitney Houston', 'de': 'Whitney Houston', 'es': 'Whitney Houston', 'nl': 'Whitney Houston', 'sv': 'Whitney Houston', 'fr': 'Whitney Houston', 'da': 'Whitney Houston'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Luxor?', 'de': 'In welchem Land liegt die antike Stadt Luxor?', 'es': '¿En que pais esta la antigua ciudad de Luxor?', 'nl': 'In welk land ligt de oude stad Luxor?', 'sv': 'I vilket land ligger den antika staden Luxor?', 'fr': "Dans quel pays se trouve l'ancienne cite de Louxor ?", 'da': 'Hvilket land huser den antikke by Luxor?'},
 'opt1': {'en': 'Sudan', 'de': 'Sudan', 'es': 'Sudan', 'nl': 'Soedan', 'sv': 'Sudan', 'fr': 'le Soudan', 'da': 'Sudan'},
 'opt2': {'en': 'Egypt', 'de': 'Ägypten', 'es': 'Egipto', 'nl': 'Egypte', 'sv': 'Egypten', 'fr': "l'Egypte", 'da': 'Egypten'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large gray mammal has a trunk and large ears, found in Africa and Asia?', 'de': 'Welches grosse graue Saeugetier hat einen Ruessel und grosse Ohren und lebt in Afrika und Asien?', 'es': '¿Que gran mamifero gris tiene trompa y orejas grandes, y se encuentra en Africa y Asia?', 'nl': 'Welk groot grijs zoogdier heeft een slurf en grote oren, gevonden in Afrika en Azie?', 'sv': 'Vilket stort gratt daggdjur har en snabel och stora oron, finns i Afrika och Asien?', 'fr': "Quel grand mammifere gris a une trompe et de grandes oreilles, present en Afrique et en Asie ?", 'da': 'Hvilket stort gråt pattedyr har en snabel og store ører, fundet i Afrika og Asien?'},
 'opt1': {'en': 'Elephant', 'de': 'Elefant', 'es': 'Elefante', 'nl': 'Olifant', 'sv': 'Elefant', 'fr': 'Elephant', 'da': 'Elefant'},
 'opt2': {'en': 'Rhinoceros', 'de': 'Nashorn', 'es': 'Rinoceronte', 'nl': 'Neushoorn', 'sv': 'Noshorning', 'fr': 'Rhinoceros', 'da': 'Næsehorn'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Indian dessert consists of fried dough balls soaked in sugar syrup?', 'de': 'Welches indische Dessert besteht aus frittierten Teigbaellchen, getraenkt in Zuckersirup?', 'es': '¿Que postre indio consiste en bolas de masa fritas remojadas en almibar?', 'nl': 'Welk Indiaas dessert bestaat uit gefrituurde deegballetjes gedrenkt in suikersiroop?', 'sv': 'Vilken indisk dessert bestar av friterade degbollar indranakta i sockersirap?', 'fr': "Quel dessert indien consiste en des boules de pate frites trempees dans un sirop de sucre ?", 'da': 'Hvilken indisk dessert består af friturestegte dejkugler dyppet i sukkersirup?'},
 'opt1': {'en': 'Gulab jamun', 'de': 'Gulab Jamun', 'es': 'Gulab jamun', 'nl': 'Gulab jamun', 'sv': 'Gulab jamun', 'fr': 'Gulab jamun', 'da': 'Gulab jamun'},
 'opt2': {'en': 'Jalebi', 'de': 'Jalebi', 'es': 'Jalebi', 'nl': 'Jalebi', 'sv': 'Jalebi', 'fr': 'Jalebi', 'da': 'Jalebi'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which First Lady was a former actress who advocated for cancer awareness after her diagnosis?', 'de': 'Welche First Lady war eine ehemalige Schauspielerin, die sich nach ihrer Diagnose fuer Krebsaufklaerung einsetzte?', 'es': '¿Que primera dama fue una ex actriz que abogo por la concienciacion sobre el cancer tras su diagnostico?', 'nl': 'Welke First Lady was een voormalige actrice die zich inzette voor kankerbewustzijn na haar diagnose?', 'sv': 'Vilken First Lady var en tidigare skadespelerska som forespakade cancermedvetenhet efter sin diagnos?', 'fr': "Quelle Première dame etait une ancienne actrice qui a plaide pour la sensibilisation au cancer apres son diagnostic ?", 'da': 'Hvilken førstedame var en tidligere skuespiller, der gik ind for kræftbevidsthed efter sin diagnose?'},
 'opt1': {'en': 'Nancy Reagan', 'de': 'Nancy Reagan', 'es': 'Nancy Reagan', 'nl': 'Nancy Reagan', 'sv': 'Nancy Reagan', 'fr': 'Nancy Reagan', 'da': 'Nancy Reagan'},
 'opt2': {'en': 'Betty Ford', 'de': 'Betty Ford', 'es': 'Betty Ford', 'nl': 'Betty Ford', 'sv': 'Betty Ford', 'fr': 'Betty Ford', 'da': 'Betty Ford'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which artist is known for the painting 'Starry Night'?", 'de': "Welcher Kuenstler ist fuer das Gemaelde 'Sternennacht' bekannt?", 'es': "¿Que artista es conocido por la pintura 'La noche estrellada'?", 'nl': "Welke kunstenaar staat bekend om het schilderij 'Sterrennacht'?", 'sv': "Vilken konstnar ar kand for malningen 'Stjarnenatt'?", 'fr': "Quel artiste est connu pour le tableau 'La Nuit etoilee' ?", 'da': "Hvilken kunstner er kendt for maleriet 'Stjernenat'?"},
 'opt1': {'en': 'Vincent van Gogh', 'de': 'Vincent van Gogh', 'es': 'Vincent van Gogh', 'nl': 'Vincent van Gogh', 'sv': 'Vincent van Gogh', 'fr': 'Vincent van Gogh', 'da': 'Vincent van Gogh'},
 'opt2': {'en': 'Paul Cezanne', 'de': 'Paul Cezanne', 'es': 'Paul Cezanne', 'nl': 'Paul Cezanne', 'sv': 'Paul Cezanne', 'fr': 'Paul Cezanne', 'da': 'Paul Cezanne'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which flowering plant is known for its fragrant white blooms and is used in perfume, native to the Philippines and Hawaii?', 'de': 'Welche bluehende Pflanze ist fuer ihre duftenden weissen Blueten bekannt und wird in Parfuem verwendet, heimisch auf den Philippinen und Hawaii?', 'es': '¿Que planta con flores es conocida por sus fragantes flores blancas y se usa en perfumeria, nativa de Filipinas y Hawai?', 'nl': 'Welke bloeiende plant staat bekend om zijn geurige witte bloesems en wordt gebruikt in parfum, inheems op de Filipijnen en Hawaii?', 'sv': 'Vilken blommande vaxt ar kand for sina doftande vita blommor och anvands i parfym, inhemsk pa Filippinerna och Hawaii?', 'fr': "Quelle plante a fleurs est connue pour ses fleurs blanches parfumees et est utilisee en parfumerie, originaire des Philippines et d'Hawaii ?", 'da': 'Hvilken blomstrende plante er kendt for sine duftende hvide blomster og bruges i parfume, hjemmehørende på Filippinerne og Hawaii?'},
 'opt1': {'en': 'Plumeria', 'de': 'Plumeria', 'es': 'Plumeria', 'nl': 'Plumeria', 'sv': 'Plumeria', 'fr': 'Plumeria (frangipanier)', 'da': 'Plumeria'},
 'opt2': {'en': 'Jasmine', 'de': 'Jasmin', 'es': 'Jazmin', 'nl': 'Jasmijn', 'sv': 'Jasmin', 'fr': 'Jasmin', 'da': 'Jasmin'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which musical features the song 'Circle of Life'?", 'de': "In welchem Musical kommt das Lied 'Circle of Life' vor?", 'es': "¿Que musical presenta la cancion 'Circle of Life'?", 'nl': "Welke musical bevat het lied 'Circle of Life'?", 'sv': "Vilken musikal innehaller latten 'Circle of Life'?", 'fr': "Quelle comedie musicale comporte la chanson 'Circle of Life' ?", 'da': "Hvilken musical har sangen 'Circle of Life'?"},
 'opt1': {'en': 'The Lion King', 'de': 'Der Koenig der Loewen', 'es': 'El Rey Leon', 'nl': 'The Lion King', 'sv': 'Lejonkungen', 'fr': 'Le Roi Lion', 'da': 'Løvernes Konge'},
 'opt2': {'en': 'Aida', 'de': 'Aida', 'es': 'Aida', 'nl': 'Aida', 'sv': 'Aida', 'fr': 'Aida', 'da': 'Aida'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country has Dublin as its capital?', 'de': 'Welches Land hat Dublin als Hauptstadt?', 'es': '¿Que pais tiene a Dublin como su capital?', 'nl': 'Welk land heeft Dublin als hoofdstad?', 'sv': 'Vilket land har Dublin som huvudstad?', 'fr': "Quel pays a Dublin pour capitale ?", 'da': 'Hvilket land har Dublin som hovedstad?'},
 'opt1': {'en': 'Scotland', 'de': 'Schottland', 'es': 'Escocia', 'nl': 'Schotland', 'sv': 'Skottland', 'fr': "l'Ecosse", 'da': 'Skotland'},
 'opt2': {'en': 'Ireland', 'de': 'Irland', 'es': 'Irlanda', 'nl': 'Ierland', 'sv': 'Irland', 'fr': "l'Irlande", 'da': 'Irland'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which spice is used to flavor traditional Christmas gingerbread alongside cinnamon and cloves?', 'de': 'Welches Gewuerz wird verwendet, um traditionellen Weihnachtslebkuchen neben Zimt und Gewuerznelken zu wuerzen?', 'es': '¿Que especia se usa para dar sabor al pan de jengibre navideño tradicional junto con la canela y el clavo?', 'nl': 'Welke specerij wordt gebruikt om traditionele kerstpeperkoek te kruiden naast kaneel en kruidnagel?', 'sv': 'Vilken krydda anvands for att smaksatta traditionell julens pepparkakor tillsammans med kanel och nejlikor?', 'fr': "Quelle epice est utilisee pour parfumer le pain d'epices de Noel traditionnel aux cotes de la cannelle et des clous de girofle ?", 'da': 'Hvilket krydderi bruges til at smage traditionel julehonningkage sammen med kanel og nelliker?'},
 'opt1': {'en': 'Ginger', 'de': 'Ingwer', 'es': 'Jengibre', 'nl': 'Gember', 'sv': 'Ingefara', 'fr': 'Le gingembre', 'da': 'Ingefær'},
 'opt2': {'en': 'Allspice', 'de': 'Piment', 'es': 'Pimienta de Jamaica', 'nl': 'Piment', 'sv': 'Kryddpeppar', 'fr': 'Le quatre-epices', 'da': 'Allehånde'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which craft involves creating jewelry by wrapping wire around beads and stones?', 'de': 'Bei welchem Handwerk wird Schmuck durch das Umwickeln von Draht um Perlen und Steine hergestellt?', 'es': '¿Que manualidad implica crear joyeria envolviendo alambre alrededor de cuentas y piedras?', 'nl': 'Bij welk handwerk maak je sieraden door draad rond kralen en stenen te wikkelen?', 'sv': 'Vilket hantverk innebar att skapa smycken genom att linda trad runt palror och stenar?', 'fr': "Quel art manuel consiste a creer des bijoux en enroulant du fil autour de perles et de pierres ?", 'da': 'Hvilket håndværk involverer at skabe smykker ved at vikle tråd om perler og sten?'},
 'opt1': {'en': 'Beading', 'de': 'Perlenarbeit', 'es': 'Abalorios', 'nl': 'Kralenwerk', 'sv': 'Palrarbete', 'fr': 'Le perlage', 'da': 'Perlearbejde'},
 'opt2': {'en': 'Wire wrapping', 'de': 'Drahtwickeltechnik', 'es': 'Enrollado de alambre', 'nl': 'Draadwikkelen', 'sv': 'Tradlindning', 'fr': 'Le wire wrapping', 'da': 'Trådvikling'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which body part contains the alveoli used for gas exchange?', 'de': 'Welcher Koerperteil enthaelt die Lungenblaeschen fuer den Gasaustausch?', 'es': '¿Que parte del cuerpo contiene los alveolos usados para el intercambio de gases?', 'nl': 'Welk lichaamsdeel bevat de longblaasjes voor gasuitwisseling?', 'sv': 'Vilken kroppsdel innehaller alveolerna som anvands for gasutbyte?', 'fr': "Quelle partie du corps contient les alveoles utilisees pour les echanges gazeux ?", 'da': 'Hvilken kropsdel indeholder alveolerne, der bruges til gasudveksling?'},
 'opt1': {'en': 'Heart', 'de': 'Herz', 'es': 'Corazon', 'nl': 'Hart', 'sv': 'Hjarta', 'fr': 'Coeur', 'da': 'Hjerte'},
 'opt2': {'en': 'Lungs', 'de': 'Lunge', 'es': 'Pulmones', 'nl': 'Longen', 'sv': 'Lungor', 'fr': 'Poumons', 'da': 'Lunger'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'In basketball, what is the term for a shot made from beyond the arc?', 'de': 'Wie nennt man beim Basketball einen Wurf, der von jenseits der Dreipunktelinie gemacht wird?', 'es': '¿En baloncesto, como se llama un lanzamiento hecho desde detras del arco?', 'nl': 'Hoe heet bij basketbal een schot van buiten de boog?', 'sv': 'Vad kallas i basket ett skott gjort utanfor bagen?', 'fr': "En basket-ball, comment appelle-t-on un tir effectue de derriere l'arc ?", 'da': 'Hvad kaldes i basketball et skud taget uden for buen?'},
 'opt1': {'en': 'Three-pointer', 'de': 'Dreipunktewurf', 'es': 'Triple', 'nl': 'Driepunter', 'sv': 'Trepoangare', 'fr': 'Un tir a trois points', 'da': 'Trepointer'},
 'opt2': {'en': 'Free throw', 'de': 'Freiwurf', 'es': 'Tiro libre', 'nl': 'Vrije worp', 'sv': 'Frikast', 'fr': 'Lancer franc', 'da': 'Straffekast'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Spanish region is known for flamenco dancing and Moorish architecture?', 'de': 'Welche spanische Region ist fuer Flamencotanz und maurische Architektur bekannt?', 'es': '¿Que region española es conocida por el baile flamenco y la arquitectura morisca?', 'nl': 'Welke Spaanse regio staat bekend om flamencodansen en Moorse architectuur?', 'sv': 'Vilken spansk region ar kand for flamencodans och morisk arkitektur?', 'fr': "Quelle region espagnole est connue pour la danse flamenco et l'architecture mauresque ?", 'da': 'Hvilken spansk region er kendt for flamencodans og maurisk arkitektur?'},
 'opt1': {'en': 'Andalusia', 'de': 'Andalusien', 'es': 'Andalucia', 'nl': 'Andalusie', 'sv': 'Andalusien', 'fr': "l'Andalousie", 'da': 'Andalusien'},
 'opt2': {'en': 'Catalonia', 'de': 'Katalonien', 'es': 'Cataluña', 'nl': 'Catalonie', 'sv': 'Katalonien', 'fr': 'la Catalogne', 'da': 'Catalonien'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which actress played Mary Poppins in the 1964 film?', 'de': 'Welche Schauspielerin spielte Mary Poppins im Film von 1964?', 'es': '¿Que actriz interpreto a Mary Poppins en la pelicula de 1964?', 'nl': 'Welke actrice speelde Mary Poppins in de film uit 1964?', 'sv': 'Vilken skadespelerska spelade Mary Poppins i 1964 ars film?', 'fr': "Quelle actrice a joue Mary Poppins dans le film de 1964 ?", 'da': 'Hvilken skuespiller spillede Mary Poppins i 1964-filmen?'},
 'opt1': {'en': 'Julie Andrews', 'de': 'Julie Andrews', 'es': 'Julie Andrews', 'nl': 'Julie Andrews', 'sv': 'Julie Andrews', 'fr': 'Julie Andrews', 'da': 'Julie Andrews'},
 'opt2': {'en': 'Angela Lansbury', 'de': 'Angela Lansbury', 'es': 'Angela Lansbury', 'nl': 'Angela Lansbury', 'sv': 'Angela Lansbury', 'fr': 'Angela Lansbury', 'da': 'Angela Lansbury'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which classic novel features the character Dorian Gray, who never ages while his portrait does?', 'de': 'In welchem Klassiker kommt die Figur Dorian Gray vor, der nie altert, waehrend sein Portrait es tut?', 'es': '¿Que novela clasica presenta al personaje Dorian Gray, que nunca envejece mientras su retrato si lo hace?', 'nl': 'In welke klassieke roman komt het personage Dorian Gray voor, die nooit ouder wordt terwijl zijn portret dat wel doet?', 'sv': 'I vilken klassisk roman forekommer karaktaren Dorian Gray, som aldrig aldras medan hans portratt gor det?', 'fr': "Quel roman classique met en scene le personnage de Dorian Gray, qui ne vieillit jamais alors que son portrait vieillit ?", 'da': 'Hvilken klassisk roman har karakteren Dorian Gray, der aldrig ældes, mens hans portræt gør?'},
 'opt1': {'en': 'Tender Is the Night', 'de': 'Zaertlich ist die Nacht', 'es': 'Suave es la noche', 'nl': 'Teder Is de Nacht', 'sv': 'Om ett ar, om en dag', 'fr': 'Tendre est la nuit', 'da': 'Blid er natten'},
 'opt2': {'en': 'The Picture of Dorian Gray', 'de': 'Das Bildnis des Dorian Gray', 'es': 'El retrato de Dorian Gray', 'nl': 'Het Portret van Dorian Gray', 'sv': 'Dorian Grays portratt', 'fr': 'Le Portrait de Dorian Gray', 'da': 'Billedet af Dorian Gray'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which 1960s band is known for 'Good Vibrations' and surf rock?", 'de': "Welche Band der 1960er ist fuer 'Good Vibrations' und Surf-Rock bekannt?", 'es': "¿Que banda de los 60 es conocida por 'Good Vibrations' y el surf rock?", 'nl': "Welke band uit de jaren 60 staat bekend om 'Good Vibrations' en surfrock?", 'sv': "Vilket 1960-talsband ar kant for 'Good Vibrations' och surfrock?", 'fr': "Quel groupe des annees 1960 est connu pour 'Good Vibrations' et le surf rock ?", 'da': "Hvilket 1960’er-band er kendt for 'Good Vibrations' og surfrock?"},
 'opt1': {'en': 'The Byrds', 'de': 'The Byrds', 'es': 'The Byrds', 'nl': 'The Byrds', 'sv': 'The Byrds', 'fr': 'The Byrds', 'da': 'The Byrds'},
 'opt2': {'en': 'The Beach Boys', 'de': 'The Beach Boys', 'es': 'The Beach Boys', 'nl': 'The Beach Boys', 'sv': 'The Beach Boys', 'fr': 'The Beach Boys', 'da': 'The Beach Boys'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a savory pie filled with meat and vegetables under a crust?', 'de': 'Welches suddstaatliche Gericht ist eine herzhafte Torte gefuellt mit Fleisch und Gemuese unter einer Kruste?', 'es': '¿Que plato sureño de EE.UU. es un pastel salado relleno de carne y verduras bajo una masa?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een hartige taart gevuld met vlees en groenten onder een korst?', 'sv': 'Vilken sydstatlig ratt ar en salt paj fylld med kott och gronsaker under en skorpa?', 'fr': "Quel plat du Sud des États-Unis est une tourte salee fourree de viande et de legumes sous une croute ?", 'da': 'Hvilken sydstatsret er en salt tærte fyldt med kød og grøntsager under en skorpe?'},
 'opt1': {'en': 'Chicken pot pie', 'de': 'Chicken Pot Pie', 'es': 'Pastel de pollo', 'nl': 'Kippenpastei', 'sv': 'Kycklingpaj', 'fr': 'La tourte au poulet', 'da': 'Kyllingetærte'},
 'opt2': {'en': "Shepherd's pie", 'de': 'Shepherd\'s Pie', 'es': 'Pastel del pastor', 'nl': "Herderspastei", 'sv': 'Herdepaj', 'fr': 'Hachis parmentier', 'da': 'Fårehyrdetærte'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which queen ruled Russia and expanded its territory significantly in the 18th century?', 'de': 'Welche Koenigin regierte Russland und erweiterte sein Territorium im 18. Jahrhundert erheblich?', 'es': '¿Que reina goberno Rusia y expandio significativamente su territorio en el siglo XVIII?', 'nl': 'Welke koningin regeerde Rusland en breidde het grondgebied aanzienlijk uit in de 18e eeuw?', 'sv': 'Vilken drottning styrde Ryssland och utokade dess territorium avsevart under 1700-talet?', 'fr': "Quelle reine a regne sur la Russie et a considerablement etendu son territoire au 18e siecle ?", 'da': 'Hvilken dronning regerede Rusland og udvidede dets territorium betydeligt i det 18. århundrede?'},
 'opt1': {'en': 'Elizabeth of Russia', 'de': 'Elisabeth von Russland', 'es': 'Isabel de Rusia', 'nl': 'Elisabeth van Rusland', 'sv': 'Elisabet av Ryssland', 'fr': 'Elisabeth de Russie', 'da': 'Elisabeth af Rusland'},
 'opt2': {'en': 'Catherine the Great', 'de': 'Katharina die Große', 'es': 'Catalina la Grande', 'nl': 'Catharina de Grote', 'sv': 'Katarina den stora', 'fr': 'Catherine la Grande', 'da': 'Katarina den Store'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which large aquatic bird has a distinctive pink color from its diet of shrimp?', 'de': 'Welcher grosse Wasservogel hat eine markante rosa Farbe von seiner Garnelendiaet?', 'es': '¿Que gran ave acuatica tiene un distintivo color rosa por su dieta de camarones?', 'nl': 'Welke grote watervogel heeft een kenmerkende roze kleur door zijn dieet van garnalen?', 'sv': 'Vilken stor vattenfagel har en markant rosa farg fran sin diet av raga?', 'fr': "Quel grand oiseau aquatique a une couleur rose distinctive due a son alimentation de crevettes ?", 'da': 'Hvilken stor vandfugl har en karakteristisk pink farve fra sin diæt af rejer?'},
 'opt1': {'en': 'Spoonbill', 'de': 'Loeffler', 'es': 'Espatula', 'nl': 'Lepelaar', 'sv': 'Skedstork', 'fr': 'Spatule', 'da': 'Skestork'},
 'opt2': {'en': 'Flamingo', 'de': 'Flamingo', 'es': 'Flamenco', 'nl': 'Flamingo', 'sv': 'Flamingo', 'fr': 'Flamant rose', 'da': 'Flamingo'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Angkor?', 'de': 'In welchem Land liegt die antike Stadt Angkor?', 'es': '¿En que pais esta la antigua ciudad de Angkor?', 'nl': 'In welk land ligt de oude stad Angkor?', 'sv': 'I vilket land ligger den antika staden Angkor?', 'fr': "Dans quel pays se trouve l'ancienne cite d'Angkor ?", 'da': 'Hvilket land huser den antikke by Angkor?'},
 'opt1': {'en': 'Laos', 'de': 'Laos', 'es': 'Laos', 'nl': 'Laos', 'sv': 'Laos', 'fr': 'le Laos', 'da': 'Laos'},
 'opt2': {'en': 'Cambodia', 'de': 'Kambodscha', 'es': 'Camboya', 'nl': 'Cambodja', 'sv': 'Kambodja', 'fr': 'le Cambodge', 'da': 'Cambodja'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of dress, popularized in the 1920s, features a straight silhouette with fringe or beading?', 'de': 'Welche Art von Kleid, populaer gemacht in den 1920er Jahren, hat eine gerade Silhouette mit Fransen oder Perlen?', 'es': '¿Que tipo de vestido, popularizado en los años 20, presenta una silueta recta con flecos o cuentas?', 'nl': 'Welk type jurk, populair gemaakt in de jaren 20, heeft een rechte silhouet met franjes of kralen?', 'sv': 'Vilken typ av klanning, popularaiserad pa 1920-talet, har en rak silhuett med fransar eller palror?', 'fr': "Quel type de robe, popularise dans les annees 1920, presente une silhouette droite avec des franges ou des perles ?", 'da': 'Hvilken type kjole, populariseret i 1920’erne, har en lige silhuet med frynser eller perler?'},
 'opt1': {'en': 'Flapper dress', 'de': 'Flapper-Kleid', 'es': 'Vestido flapper', 'nl': 'Flapperjurk', 'sv': 'Flapperklanning', 'fr': 'Robe charleston', 'da': 'Flapper-kjole'},
 'opt2': {'en': 'Empire waist dress', 'de': 'Empire-Kleid', 'es': 'Vestido de talle imperio', 'nl': 'Empirejurk', 'sv': 'Empirklanning', 'fr': 'Robe taille empire', 'da': 'Empire-kjole'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which streaming series follows a group of friends navigating dating in New York, a modern take on a 1990s classic?', 'de': 'Welche Streaming-Serie folgt einer Gruppe von Freunden, die in New York daten, eine moderne Version eines Klassikers der 1990er?', 'es': '¿Que serie de streaming sigue a un grupo de amigas navegando por las citas en Nueva York, una version moderna de un clasico de los 90?', 'nl': 'Welke streamingserie volgt een groep vrienden die daten in New York, een moderne versie van een klassieker uit de jaren 90?', 'sv': 'Vilken streamingserie foljer en grupp vanner som dejtar i New York, en modern tolkning av en 1990-talsklassiker?', 'fr': "Quelle serie en streaming suit un groupe d'amies naviguant dans les rencontres a New York, une version moderne d'un classique des annees 1990 ?", 'da': 'Hvilken streamingserie følger en gruppe venner, der navigerer i datinglivet i New York, en moderne genfortolkning af en 1990’er-klassiker?'},
 'opt1': {'en': 'And Just Like That', 'de': 'And Just Like That', 'es': 'And Just Like That', 'nl': 'And Just Like That', 'sv': 'And Just Like That', 'fr': 'And Just Like That', 'da': 'And Just Like That'},
 'opt2': {'en': 'Girls', 'de': 'Girls', 'es': 'Girls', 'nl': 'Girls', 'sv': 'Girls', 'fr': 'Girls', 'da': 'Girls'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which tree produces the seeds used to make chocolate?', 'de': 'Welcher Baum produziert die Samen, aus denen Schokolade hergestellt wird?', 'es': '¿Que arbol produce las semillas usadas para hacer chocolate?', 'nl': 'Welke boom produceert de zaden die worden gebruikt om chocolade te maken?', 'sv': 'Vilket trad producerar fronen som anvands for att gora choklad?', 'fr': "Quel arbre produit les graines utilisees pour fabriquer le chocolat ?", 'da': 'Hvilket træ producerer de frø, der bruges til at lave chokolade?'},
 'opt1': {'en': 'Cacao tree', 'de': 'Kakaobaum', 'es': 'Arbol de cacao', 'nl': 'Cacaoboom', 'sv': 'Kakaotrad', 'fr': 'Cacaoyer', 'da': 'Kakaotræ'},
 'opt2': {'en': 'Coffee tree', 'de': 'Kaffeebaum', 'es': 'Cafeto', 'nl': 'Koffieboom', 'sv': 'Kaffetrad', 'fr': 'Cafeier', 'da': 'Kaffetræ'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Korean dish is a spicy fermented vegetable side dish?', 'de': 'Welches koreanische Gericht ist eine scharfe, fermentierte Gemuesebeilage?', 'es': '¿Que plato coreano es un acompañamiento de verduras fermentadas y picantes?', 'nl': 'Welk Koreaans gerecht is een pittig gefermenteerd groentebijgerecht?', 'sv': 'Vilken koreansk ratt ar ett kryddigt fermenterat gronsakstillbehor?', 'fr': "Quel plat coreen est un accompagnement de legumes fermentes et epices ?", 'da': 'Hvilken koreansk ret er et krydret fermenteret grøntsagstilbehør?'},
 'opt1': {'en': 'Banchan', 'de': 'Banchan', 'es': 'Banchan', 'nl': 'Banchan', 'sv': 'Banchan', 'fr': 'Banchan', 'da': 'Banchan'},
 'opt2': {'en': 'Kimchi', 'de': 'Kimchi', 'es': 'Kimchi', 'nl': 'Kimchi', 'sv': 'Kimchi', 'fr': 'Kimchi', 'da': 'Kimchi'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which bone protects the brain?', 'de': 'Welcher Knochen schuetzt das Gehirn?', 'es': '¿Que hueso protege el cerebro?', 'nl': 'Welk bot beschermt de hersenen?', 'sv': 'Vilket ben skyddar hjarnan?', 'fr': "Quel os protege le cerveau ?", 'da': 'Hvilken knogle beskytter hjernen?'},
 'opt1': {'en': 'Spine', 'de': 'Wirbelsaeule', 'es': 'Columna vertebral', 'nl': 'Ruggengraat', 'sv': 'Ryggrad', 'fr': 'La colonne vertebrale', 'da': 'Rygsøjle'},
 'opt2': {'en': 'Skull', 'de': 'Schaedel', 'es': 'Craneo', 'nl': 'Schedel', 'sv': 'Skalle', 'fr': 'Le crane', 'da': 'Kranie'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Palenque?', 'de': 'In welchem Land liegt die antike Stadt Palenque?', 'es': '¿En que pais esta la antigua ciudad de Palenque?', 'nl': 'In welk land ligt de oude stad Palenque?', 'sv': 'I vilket land ligger den antika staden Palenque?', 'fr': "Dans quel pays se trouve l'ancienne cite de Palenque ?", 'da': 'Hvilket land huser den antikke by Palenque?'},
 'opt1': {'en': 'Guatemala', 'de': 'Guatemala', 'es': 'Guatemala', 'nl': 'Guatemala', 'sv': 'Guatemala', 'fr': 'le Guatemala', 'da': 'Guatemala'},
 'opt2': {'en': 'Mexico', 'de': 'Mexiko', 'es': 'Mexico', 'nl': 'Mexico', 'sv': 'Mexiko', 'fr': 'le Mexique', 'da': 'Mexico'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which First Lady chaired the U.N. Commission on Human Rights after her husband's presidency?", 'de': 'Welche First Lady leitete die UN-Menschenrechtskommission nach der Praesidentschaft ihres Mannes?', 'es': '¿Que primera dama presidio la Comision de Derechos Humanos de la ONU tras la presidencia de su esposo?', 'nl': 'Welke First Lady leidde de VN-Mensenrechtencommissie na het presidentschap van haar man?', 'sv': 'Vilken First Lady ledde FN:s manniskorattskommission efter sin mans presidentskap?', 'fr': "Quelle Première dame a preside la Commission des droits de l'homme de l'ONU apres la presidence de son mari ?", 'da': 'Hvilken førstedame ledede FN’s Menneskerettighedskommission efter sin mands præsidentskab?'},
 'opt1': {'en': 'Grace Coolidge', 'de': 'Grace Coolidge', 'es': 'Grace Coolidge', 'nl': 'Grace Coolidge', 'sv': 'Grace Coolidge', 'fr': 'Grace Coolidge', 'da': 'Grace Coolidge'},
 'opt2': {'en': 'Eleanor Roosevelt', 'de': 'Eleanor Roosevelt', 'es': 'Eleanor Roosevelt', 'nl': 'Eleanor Roosevelt', 'sv': 'Eleanor Roosevelt', 'fr': 'Eleanor Roosevelt', 'da': 'Eleanor Roosevelt'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which composer wrote 'The Rite of Spring', causing a riot at its premiere?", 'de': "Welcher Komponist schrieb 'Le Sacre du Printemps', was bei der Urauffuehrung einen Aufruhr verursachte?", 'es': "¿Que compositor escribio 'La consagracion de la primavera', causando un disturbio en su estreno?", 'nl': "Welke componist schreef 'Le Sacre du Printemps', wat een rel veroorzaakte bij de premiere?", 'sv': "Vilken kompositor skrev 'Varens offer', som orsakade upplopp vid premiaren?", 'fr': "Quel compositeur a ecrit 'Le Sacre du printemps', provoquant une emeute lors de sa premiere ?", 'da': "Hvilken komponist skrev 'Le Sacre du Printemps', der forårsagede optøjer ved premieren?"},
 'opt1': {'en': 'Igor Stravinsky', 'de': 'Igor Strawinsky', 'es': 'Igor Stravinski', 'nl': 'Igor Stravinsky', 'sv': 'Igor Stravinskij', 'fr': 'Igor Stravinski', 'da': 'Igor Stravinskij'},
 'opt2': {'en': 'Sergei Rachmaninoff', 'de': 'Sergei Rachmaninow', 'es': 'Sergei Rajmaninov', 'nl': 'Sergej Rachmaninov', 'sv': 'Sergej Rachmaninov', 'fr': 'Sergueï Rachmaninov', 'da': 'Sergej Rachmaninov'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Truvy in the film 'Steel Magnolias'?", 'de': "Welche Schauspielerin spielte Truvy im Film 'Stahlmagnolien'?", 'es': "¿Que actriz interpreto a Truvy en la pelicula 'Magnolias de acero'?", 'nl': "Welke actrice speelde Truvy in de film 'Steel Magnolias'?", 'sv': "Vilken skadespelerska spelade Truvy i filmen 'Stalmagnolior'?", 'fr': "Quelle actrice a joue Truvy dans le film 'Potins de femmes' ?", 'da': "Hvilken skuespiller spillede Truvy i filmen 'Steel Magnolias'?"},
 'opt1': {'en': 'Dolly Parton', 'de': 'Dolly Parton', 'es': 'Dolly Parton', 'nl': 'Dolly Parton', 'sv': 'Dolly Parton', 'fr': 'Dolly Parton', 'da': 'Dolly Parton'},
 'opt2': {'en': 'Sally Field', 'de': 'Sally Field', 'es': 'Sally Field', 'nl': 'Sally Field', 'sv': 'Sally Field', 'fr': 'Sally Field', 'da': 'Sally Field'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Southern U.S. dish is deep-fried green tomato slices, breaded and crispy?', 'de': 'Welches suddstaatliche Gericht sind frittierte gruene Tomatenscheiben, paniert und knusprig?', 'es': '¿Que plato sureño de EE.UU. son rodajas de tomate verde fritas, empanadas y crujientes?', 'nl': 'Welk zuidelijk Amerikaans gerecht is gefrituurde groene tomatenschijfjes, gepaneerd en krokant?', 'sv': 'Vilken sydstatlig ratt ar friterade grona tomatskivor, panerade och krispiga?', 'fr': "Quel plat du Sud des États-Unis est des tranches de tomates vertes frites, panees et croustillantes ?", 'da': 'Hvilken sydstatsret er friturestegte grønne tomatskiver, paneret og sprøde?'},
 'opt1': {'en': 'Fried okra', 'de': 'Frittierte Okra', 'es': 'Okra frita', 'nl': 'Gefrituurde okra', 'sv': 'Friterad okra', 'fr': 'Gombo frit', 'da': 'Friturestegt okra'},
 'opt2': {'en': 'Fried green tomatoes', 'de': 'Fried Green Tomatoes', 'es': 'Tomates verdes fritos', 'nl': 'Gefrituurde groene tomaten', 'sv': 'Friterade grona tomater', 'fr': 'Tomates vertes frites', 'da': 'Friturestegte grønne tomater'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which U.S. state is home to the Blue Ridge Mountains?', 'de': 'In welchem US-Bundesstaat liegen die Blue Ridge Mountains?', 'es': '¿En que estado de EE.UU. estan las montañas Blue Ridge?', 'nl': 'In welke Amerikaanse staat liggen de Blue Ridge Mountains?', 'sv': 'I vilken amerikansk delstat ligger Blue Ridge Mountains?', 'fr': "Dans quel État americain se trouvent les Blue Ridge Mountains ?", 'da': 'I hvilken amerikansk delstat ligger Blue Ridge-bjergene?'},
 'opt1': {'en': 'West Virginia', 'de': 'West Virginia', 'es': 'Virginia Occidental', 'nl': 'West Virginia', 'sv': 'West Virginia', 'fr': 'la Virginie-Occidentale', 'da': 'West Virginia'},
 'opt2': {'en': 'Virginia', 'de': 'Virginia', 'es': 'Virginia', 'nl': 'Virginia', 'sv': 'Virginia', 'fr': 'la Virginie', 'da': 'Virginia'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which small mammal is known for its distinctive black and white striped body and spraying odor?', 'de': 'Welches kleine Saeugetier ist fuer seinen markanten schwarz-weiss gestreiften Koerper und seinen Sprühgeruch bekannt?', 'es': '¿Que pequeño mamifero es conocido por su distintivo cuerpo rayado en blanco y negro y su olor rociado?', 'nl': 'Welk klein zoogdier staat bekend om zijn kenmerkende zwart-wit gestreepte lichaam en spuitende geur?', 'sv': 'Vilket litet daggdjur ar kant for sin markanta svartvitrandiga kropp och spraylukt?', 'fr': "Quel petit mammifere est connu pour son corps distinctif raye de noir et blanc et son odeur pulverisee ?", 'da': 'Hvilket lille pattedyr er kendt for sin karakteristiske sort-hvide stribede krop og sprøjtende lugt?'},
 'opt1': {'en': 'Skunk', 'de': 'Stinktier', 'es': 'Mofeta', 'nl': 'Stinkdier', 'sv': 'Skunk', 'fr': 'Mouffette', 'da': 'Stinkdyr'},
 'opt2': {'en': 'Badger', 'de': 'Dachs', 'es': 'Tejon', 'nl': 'Das', 'sv': 'Gravling', 'fr': 'Blaireau', 'da': 'Grævling'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Indian tradition involves applying henna designs to hands before weddings?', 'de': 'Welche indische Tradition beinhaltet das Auftragen von Henna-Mustern auf die Haende vor Hochzeiten?', 'es': '¿Que tradicion india implica aplicar diseños de henna en las manos antes de las bodas?', 'nl': 'Welke Indiase traditie omvat het aanbrengen van hennapatronen op handen voor huwelijken?', 'sv': 'Vilken indisk tradition innebar att applicera hennamonster pa hander fore brollop?', 'fr': "Quelle tradition indienne implique d'appliquer des motifs au henne sur les mains avant les mariages ?", 'da': 'Hvilken indisk tradition involverer at påføre hennamønstre på hænder før bryllupper?'},
 'opt1': {'en': 'Mehndi', 'de': 'Mehndi', 'es': 'Mehndi', 'nl': 'Mehndi', 'sv': 'Mehndi', 'fr': 'Mehndi', 'da': 'Mehndi'},
 'opt2': {'en': 'Rangoli', 'de': 'Rangoli', 'es': 'Rangoli', 'nl': 'Rangoli', 'sv': 'Rangoli', 'fr': 'Rangoli', 'da': 'Rangoli'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which author wrote 'The Kite Runner'?", 'de': "Welcher Autor schrieb 'Drachenlaeufer'?", 'es': "¿Que autor escribio 'Cometas en el cielo'?", 'nl': "Welke auteur schreef 'De Vliegeraar'?", 'sv': "Vilken forfattare skrev 'Flyktingpojken'?", 'fr': "Quel auteur a ecrit 'Les Cerfs-volants de Kaboul' ?", 'da': "Hvilken forfatter skrev 'Drageløberen'?"},
 'opt1': {'en': 'Khaled Hosseini', 'de': 'Khaled Hosseini', 'es': 'Khaled Hosseini', 'nl': 'Khaled Hosseini', 'sv': 'Khaled Hosseini', 'fr': 'Khaled Hosseini', 'da': 'Khaled Hosseini'},
 'opt2': {'en': 'Amy Tan', 'de': 'Amy Tan', 'es': 'Amy Tan', 'nl': 'Amy Tan', 'sv': 'Amy Tan', 'fr': 'Amy Tan', 'da': 'Amy Tan'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which country is home to the ancient city of Petra's rival tourist site, Wadi Rum desert?", 'de': 'In welchem Land liegt die Wueste Wadi Rum, ein Konkurrent von Petra als Touristenziel?', 'es': '¿En que pais esta el desierto de Wadi Rum, sitio turistico rival de Petra?', 'nl': 'In welk land ligt de Wadi Rum-woestijn, een rivaliserende toeristische bestemming van Petra?', 'sv': 'I vilket land ligger Wadi Rum-oknen, Petras rivaliserande turistmal?', 'fr': "Dans quel pays se trouve le desert du Wadi Rum, site touristique rival de Petra ?", 'da': 'Hvilket land huser Wadi Rum-ørkenen, Petras rivaliserende turiststed?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Egypt', 'de': 'Ägypten', 'es': 'Egipto', 'nl': 'Egypte', 'sv': 'Egypten', 'fr': "l'Egypte", 'da': 'Egypten'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which instrument is central to Hawaiian music, a small four-stringed guitar?', 'de': 'Welches Instrument ist zentral fuer hawaiianische Musik, eine kleine viersaitige Gitarre?', 'es': '¿Que instrumento es central en la musica hawaiana, una pequeña guitarra de cuatro cuerdas?', 'nl': 'Welk instrument is centraal in Hawaiiaanse muziek, een kleine viersnarige gitaar?', 'sv': 'Vilket instrument ar centralt i hawaiiansk musik, en liten fyrstrangad gitarr?', 'fr': "Quel instrument est central dans la musique hawaiienne, une petite guitare a quatre cordes ?", 'da': 'Hvilket instrument er centralt i hawaiiansk musik, en lille firestrenget guitar?'},
 'opt1': {'en': 'Mandolin', 'de': 'Mandoline', 'es': 'Mandolina', 'nl': 'Mandoline', 'sv': 'Mandolin', 'fr': 'Mandoline', 'da': 'Mandolin'},
 'opt2': {'en': 'Ukulele', 'de': 'Ukulele', 'es': 'Ukelele', 'nl': 'Ukelele', 'sv': 'Ukulele', 'fr': 'Ukulele', 'da': 'Ukulele'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a cornbread-based dressing served at Thanksgiving?', 'de': 'Welches suddstaatliche Gericht ist eine Fuellung auf Maisbrotbasis, serviert zu Thanksgiving?', 'es': '¿Que plato sureño de EE.UU. es un relleno a base de pan de maiz servido en Accion de Gracias?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een vulling op basis van maisbrood, geserveerd met Thanksgiving?', 'sv': 'Vilken sydstatlig ratt ar en fyllning baserad pa majsbrod, serverad pa Thanksgiving?', 'fr': "Quel plat du Sud des États-Unis est une farce a base de pain de mais servie a Thanksgiving ?", 'da': 'Hvilken sydstatsret er en fyld baseret på majsbrød, serveret til Thanksgiving?'},
 'opt1': {'en': 'Cornbread dressing', 'de': 'Maisbrot-Fuellung', 'es': 'Relleno de pan de maiz', 'nl': 'Maisbroodvulling', 'sv': 'Majsbrödsfyllning', 'fr': 'Farce au pain de mais', 'da': 'Majsbrødfyld'},
 'opt2': {'en': 'Bread stuffing', 'de': 'Brotfuellung', 'es': 'Relleno de pan', 'nl': 'Broodvulling', 'sv': 'Brödfyllning', 'fr': 'Farce au pain', 'da': 'Brødfyld'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which mineral, found in dairy, strengthens bones and teeth?', 'de': 'Welches Mineral, in Milchprodukten enthalten, staerkt Knochen und Zaehne?', 'es': '¿Que mineral, presente en lacteos, fortalece huesos y dientes?', 'nl': 'Welk mineraal, aanwezig in zuivel, versterkt botten en tanden?', 'sv': 'Vilket mineral, som finns i mejeriprodukter, starker ben och tander?', 'fr': "Quel mineral, present dans les produits laitiers, renforce les os et les dents ?", 'da': 'Hvilket mineral, fundet i mejeriprodukter, styrker knogler og tænder?'},
 'opt1': {'en': 'Calcium', 'de': 'Kalzium', 'es': 'Calcio', 'nl': 'Calcium', 'sv': 'Kalcium', 'fr': 'Calcium', 'da': 'Calcium'},
 'opt2': {'en': 'Phosphorus', 'de': 'Phosphor', 'es': 'Fosforo', 'nl': 'Fosfor', 'sv': 'Fosfor', 'fr': 'Phosphore', 'da': 'Fosfor'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which film has Meryl Streep playing Julia Child in a dual-timeline story?', 'de': 'In welchem Film spielt Meryl Streep Julia Child in einer Geschichte mit zwei Zeitebenen?', 'es': '¿En que pelicula Meryl Streep interpreta a Julia Child en una historia de dos lineas temporales?', 'nl': 'In welke film speelt Meryl Streep Julia Child in een verhaal met twee tijdlijnen?', 'sv': 'I vilken film spelar Meryl Streep Julia Child i en historia med tva tidslinjer?', 'fr': "Dans quel film Meryl Streep joue-t-elle Julia Child dans une histoire a double chronologie ?", 'da': 'I hvilken film spiller Meryl Streep Julia Child i en historie med to tidslinjer?'},
 'opt1': {'en': 'Julie & Julia', 'de': 'Julie & Julia', 'es': 'Julie & Julia', 'nl': 'Julie & Julia', 'sv': 'Julie & Julia', 'fr': 'Julie & Julia', 'da': 'Julie & Julia'},
 'opt2': {'en': 'Julia', 'de': 'Julia', 'es': 'Julia', 'nl': 'Julia', 'sv': 'Julia', 'fr': 'Julia', 'da': 'Julia'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which country is home to the ancient city of Petra's neighbor, Amman, its capital?", 'de': 'In welchem Land liegt Amman, die Hauptstadt in der Naehe von Petra?', 'es': '¿En que pais esta Amman, la capital cercana a Petra?', 'nl': 'In welk land ligt Amman, de hoofdstad nabij Petra?', 'sv': 'I vilket land ligger Amman, huvudstaden nara Petra?', 'fr': "Dans quel pays se trouve Amman, la capitale voisine de Petra ?", 'da': 'Hvilket land huser Amman, hovedstaden nær Petra?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Lebanon', 'de': 'Libanon', 'es': 'Libano', 'nl': 'Libanon', 'sv': 'Libanon', 'fr': 'le Liban', 'da': 'Libanon'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of pants, popularized in the 1960s-70s, flare out below the knee?', 'de': 'Welche Art von Hose, populaer gemacht in den 1960er-70er Jahren, weitet sich unterhalb des Knies aus?', 'es': '¿Que tipo de pantalones, popularizados en los años 60-70, se ensanchan debajo de la rodilla?', 'nl': 'Welk type broek, populair gemaakt in de jaren 60-70, wordt breder onder de knie?', 'sv': 'Vilken typ av byxor, popularaiserade pa 1960-70-talen, blir vidare under knat?', 'fr': "Quel type de pantalon, popularise dans les annees 1960-70, s'evase sous le genou ?", 'da': 'Hvilken type bukser, populariseret i 1960’erne-70’erne, flarer ud under knæet?'},
 'opt1': {'en': 'Bell-bottoms', 'de': 'Schlaghosen', 'es': 'Pantalones acampanados', 'nl': 'Broek met wijde pijpen', 'sv': 'Flarbyxor', 'fr': 'Pantalon patte d\'eph', 'da': 'Klokkebukser'},
 'opt2': {'en': 'Skinny jeans', 'de': 'Skinny Jeans', 'es': 'Jeans pitillo', 'nl': 'Skinny jeans', 'sv': 'Skinny jeans', 'fr': 'Jean slim', 'da': 'Skinny jeans'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which large bird of prey has a distinctive white head and is a U.S. national symbol?', 'de': 'Welcher grosse Greifvogel hat einen markanten weissen Kopf und ist ein Nationalsymbol der USA?', 'es': '¿Que gran ave rapaz tiene una distintiva cabeza blanca y es simbolo nacional de EE.UU.?', 'nl': 'Welke grote roofvogel heeft een kenmerkende witte kop en is een Amerikaans nationaal symbool?', 'sv': 'Vilken stor rovfagel har ett markant vitt huvud och ar en amerikansk nationalsymbol?', 'fr': "Quel grand rapace a une tete blanche distinctive et est un symbole national americain ?", 'da': 'Hvilken stor rovfugl har et karakteristisk hvidt hoved og er et amerikansk nationalsymbol?'},
 'opt1': {'en': 'Osprey', 'de': 'Fischadler', 'es': 'Aguila pescadora', 'nl': 'Visarend', 'sv': 'Fiskgjuse', 'fr': 'Balbuzard', 'da': 'Fiskeørn'},
 'opt2': {'en': 'Bald eagle', 'de': 'Weisskopfseeadler', 'es': 'Aguila calva', 'nl': 'Amerikaanse zeearend', 'sv': 'Vithuvad havsorn', 'fr': "Aigle a tete blanche", 'da': 'Amerikansk havørn'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which First Lady was known for restoring the White House's historic furnishings in the early 1960s?", 'de': "Welche First Lady war bekannt fuer die Restaurierung der historischen Einrichtung des Weissen Hauses in den fruehen 1960er Jahren?", 'es': "¿Que primera dama fue conocida por restaurar el mobiliario historico de la Casa Blanca a principios de los años 60?", 'nl': "Welke First Lady stond bekend om het restaureren van het historisch meubilair van het Witte Huis begin jaren 60?", 'sv': "Vilken First Lady var kand for att restaurera Vita husets historiska inredning i borjan av 1960-talet?", 'fr': "Quelle Première dame etait connue pour avoir restaure le mobilier historique de la Maison-Blanche au debut des annees 1960 ?", 'da': 'Hvilken førstedame var kendt for at genoprette Det Hvide Hus’ historiske inventar i begyndelsen af 1960’erne?'},
 'opt1': {'en': 'Pat Nixon', 'de': 'Pat Nixon', 'es': 'Pat Nixon', 'nl': 'Pat Nixon', 'sv': 'Pat Nixon', 'fr': 'Pat Nixon', 'da': 'Pat Nixon'},
 'opt2': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'}},

{'category': 'History', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which First Lady taught at a school for the deaf before her husband became president?", 'de': 'Welche First Lady unterrichtete an einer Gehoerlosenschule, bevor ihr Mann Praesident wurde?', 'es': '¿Que primera dama enseño en una escuela para sordos antes de que su esposo fuera presidente?', 'nl': 'Welke First Lady gaf les op een school voor doven voordat haar man president werd?', 'sv': 'Vilken First Lady undervisade pa en dovskola innan hennes man blev president?', 'fr': "Quelle Première dame a enseigne dans une ecole pour sourds avant que son mari ne devienne president ?", 'da': 'Hvilken førstedame underviste på en skole for døve, før hendes mand blev præsident?'},
 'opt1': {'en': 'Bess Truman', 'de': 'Bess Truman', 'es': 'Bess Truman', 'nl': 'Bess Truman', 'sv': 'Bess Truman', 'fr': 'Bess Truman', 'da': 'Bess Truman'},
 'opt2': {'en': 'Grace Coolidge', 'de': 'Grace Coolidge', 'es': 'Grace Coolidge', 'nl': 'Grace Coolidge', 'sv': 'Grace Coolidge', 'fr': 'Grace Coolidge', 'da': 'Grace Coolidge'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which Indian author wrote 'The God of Small Things'?", 'de': "Welche indische Autorin schrieb 'Der Gott der kleinen Dinge'?", 'es': "¿Que autora india escribio 'El dios de las pequeñas cosas'?", 'nl': "Welke Indiase auteur schreef 'De God van Kleine Dingen'?", 'sv': "Vilken indisk forfattare skrev 'De sma tingens gud'?", 'fr': "Quelle autrice indienne a ecrit 'Le Dieu des Petits Riens' ?", 'da': "Hvilken indisk forfatter skrev 'De små tings gud'?"},
 'opt1': {'en': 'Jhumpa Lahiri', 'de': 'Jhumpa Lahiri', 'es': 'Jhumpa Lahiri', 'nl': 'Jhumpa Lahiri', 'sv': 'Jhumpa Lahiri', 'fr': 'Jhumpa Lahiri', 'da': 'Jhumpa Lahiri'},
 'opt2': {'en': 'Arundhati Roy', 'de': 'Arundhati Roy', 'es': 'Arundhati Roy', 'nl': 'Arundhati Roy', 'sv': 'Arundhati Roy', 'fr': 'Arundhati Roy', 'da': 'Arundhati Roy'}},

{'category': 'Science', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which vitamin, found in eggs and liver, supports healthy vision and immune function?', 'de': 'Welches Vitamin, in Eiern und Leber enthalten, unterstuetzt gesundes Sehen und die Immunfunktion?', 'es': '¿Que vitamina, presente en huevos e higado, apoya la vision saludable y la funcion inmune?', 'nl': 'Welk vitamine, aanwezig in eieren en lever, ondersteunt gezond zicht en immuunfunctie?', 'sv': 'Vilket vitamin, som finns i agg och lever, stodjer god syn och immunfunktion?', 'fr': "Quelle vitamine, presente dans les oeufs et le foie, soutient une bonne vision et la fonction immunitaire ?", 'da': 'Hvilket vitamin, fundet i æg og lever, understøtter sundt syn og immunfunktion?'},
 'opt1': {'en': 'Vitamin E', 'de': 'Vitamin E', 'es': 'Vitamina E', 'nl': 'Vitamine E', 'sv': 'Vitamin E', 'fr': 'Vitamine E', 'da': 'Vitamin E'},
 'opt2': {'en': 'Vitamin A', 'de': 'Vitamin A', 'es': 'Vitamina A', 'nl': 'Vitamine A', 'sv': 'Vitamin A', 'fr': 'Vitamine A', 'da': 'Vitamin A'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Italian sculptor and painter created the ceiling of the Sistine Chapel?', 'de': 'Welcher italienische Bildhauer und Maler schuf die Decke der Sixtinischen Kapelle?', 'es': '¿Que escultor y pintor italiano creo el techo de la Capilla Sixtina?', 'nl': 'Welke Italiaanse beeldhouwer en schilder creëerde het plafond van de Sixtijnse Kapel?', 'sv': 'Vilken italiensk skulptor och malare skapade taket i Sixtinska kapellet?', 'fr': "Quel sculpteur et peintre italien a cree le plafond de la chapelle Sixtine ?", 'da': 'Hvilken italiensk billedhugger og maler skabte loftet i Det Sixtinske Kapel?'},
 'opt1': {'en': 'Michelangelo', 'de': 'Michelangelo', 'es': 'Miguel Angel', 'nl': 'Michelangelo', 'sv': 'Michelangelo', 'fr': 'Michel-Ange', 'da': 'Michelangelo'},
 'opt2': {'en': 'Raphael', 'de': 'Raffael', 'es': 'Rafael', 'nl': 'Rafael', 'sv': 'Rafael', 'fr': 'Raphael', 'da': 'Rafael'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which South American country is home to the Atacama Desert, the driest place on Earth?', 'de': 'In welchem suedamerikanischen Land liegt die Atacamawueste, der trockenste Ort der Erde?', 'es': '¿Que pais sudamericano alberga el Desierto de Atacama, el lugar mas seco de la Tierra?', 'nl': 'Welk Zuid-Amerikaans land herbergt de Atacamawoestijn, de droogste plek op aarde?', 'sv': 'Vilket sydamerikanskt land huserar Atacamaoknen, den torraste platsen pa jorden?', 'fr': "Quel pays sud-americain abrite le desert d'Atacama, l'endroit le plus sec de la Terre ?", 'da': 'Hvilket sydamerikansk land huser Atacama-ørkenen, det tørreste sted på Jorden?'},
 'opt1': {'en': 'Chile', 'de': 'Chile', 'es': 'Chile', 'nl': 'Chili', 'sv': 'Chile', 'fr': 'le Chili', 'da': 'Chile'},
 'opt2': {'en': 'Peru', 'de': 'Peru', 'es': 'Peru', 'nl': 'Peru', 'sv': 'Peru', 'fr': 'le Perou', 'da': 'Peru'}},

{'category': 'Music', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Russian composer wrote 'Swan Lake' and the '1812 Overture'?", 'de': "Welcher russische Komponist schrieb 'Schwanensee' und die 'Ouvertuere 1812'?", 'es': "¿Que compositor ruso escribio 'El lago de los cisnes' y la 'Obertura 1812'?", 'nl': "Welke Russische componist schreef 'Het Zwanenmeer' en het 'Ouverture van 1812'?", 'sv': "Vilken rysk kompositor skrev 'Svansjon' och '1812-overturen'?", 'fr': "Quel compositeur russe a ecrit 'Le Lac des cygnes' et l''Ouverture 1812' ?", 'da': "Hvilken russisk komponist skrev 'Svanesøen' og '1812-ouverturen'?"},
 'opt1': {'en': 'Pyotr Ilyich Tchaikovsky', 'de': 'Pjotr Iljitsch Tschaikowski', 'es': 'Piotr Ilich Chaikovski', 'nl': 'Pjotr Iljitsj Tsjaikovski', 'sv': 'Pjotr Iljitj Tjajkovskij', 'fr': 'Piotr Ilitch Tchaikovski', 'da': 'Pjotr Iljitj Tjajkovskij'},
 'opt2': {'en': 'Modest Mussorgsky', 'de': 'Modest Mussorgski', 'es': 'Modest Musorgski', 'nl': 'Modest Moessorgski', 'sv': 'Modest Musorgskij', 'fr': 'Modeste Moussorgski', 'da': 'Modest Musorgskij'}},

{'category': 'Food', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which French dessert is a light, airy baked dish that can collapse if opened too early?', 'de': 'Welches franzoesische Dessert ist ein leichtes, luftiges Backwerk, das zusammenfallen kann, wenn es zu frueh geoeffnet wird?', 'es': '¿Que postre frances es un plato horneado ligero y aereado que puede colapsar si se abre demasiado pronto?', 'nl': 'Welk Frans dessert is een licht, luchtig gebakken gerecht dat kan inzakken als het te vroeg wordt geopend?', 'sv': 'Vilken fransk dessert ar en latt, luftig ugnsbakad ratt som kan falla ihop om den oppnas for tidigt?', 'fr': "Quel dessert français est un plat cuit au four leger et aere qui peut s'effondrer s'il est ouvert trop tot ?", 'da': 'Hvilken fransk dessert er en let, luftig bagt ret, der kan falde sammen, hvis den åbnes for tidligt?'},
 'opt1': {'en': 'Creme brulee', 'de': 'Creme brulee', 'es': 'Creme brulee', 'nl': 'Creme brulee', 'sv': 'Creme brulee', 'fr': 'Creme brulee', 'da': 'Creme brulee'},
 'opt2': {'en': 'Souffle', 'de': 'Souffle', 'es': 'Souffle', 'nl': 'Souffle', 'sv': 'Souffle', 'fr': 'Souffle', 'da': 'Souffle'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Greek mythology, which Titan is condemned to hold up the sky?', 'de': 'In der griechischen Mythologie, welcher Titan ist dazu verdammt, den Himmel zu tragen?', 'es': '¿En la mitologia griega, que titan esta condenado a sostener el cielo?', 'nl': 'In de Griekse mythologie, welke Titaan is veroordeeld om de hemel te dragen?', 'sv': 'I grekisk mytologi, vilken titan ar domd att halla upp himlen?', 'fr': "Dans la mythologie grecque, quel Titan est condamne a porter le ciel ?", 'da': 'I græsk mytologi, hvilken titan er dømt til at holde himlen oppe?'},
 'opt1': {'en': 'Atlas', 'de': 'Atlas', 'es': 'Atlas', 'nl': 'Atlas', 'sv': 'Atlas', 'fr': 'Atlas', 'da': 'Atlas'},
 'opt2': {'en': 'Prometheus', 'de': 'Prometheus', 'es': 'Prometeo', 'nl': 'Prometheus', 'sv': 'Prometheus', 'fr': 'Promethee', 'da': 'Prometheus'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which actress won an Oscar for playing Virginia Woolf in 'The Hours'?", 'de': "Welche Schauspielerin gewann einen Oscar fuer ihre Rolle als Virginia Woolf in 'The Hours'?", 'es': "¿Que actriz gano un Oscar por interpretar a Virginia Woolf en 'Las horas'?", 'nl': "Welke actrice won een Oscar voor haar rol als Virginia Woolf in 'The Hours'?", 'sv': "Vilken skadespelerska vann en Oscar for att ha spelat Virginia Woolf i 'The Hours'?", 'fr': "Quelle actrice a remporte un Oscar pour son role de Virginia Woolf dans 'The Hours' ?", 'da': "Hvilken skuespiller vandt en Oscar for at spille Virginia Woolf i 'The Hours'?"},
 'opt1': {'en': 'Julianne Moore', 'de': 'Julianne Moore', 'es': 'Julianne Moore', 'nl': 'Julianne Moore', 'sv': 'Julianne Moore', 'fr': 'Julianne Moore', 'da': 'Julianne Moore'},
 'opt2': {'en': 'Nicole Kidman', 'de': 'Nicole Kidman', 'es': 'Nicole Kidman', 'nl': 'Nicole Kidman', 'sv': 'Nicole Kidman', 'fr': 'Nicole Kidman', 'da': 'Nicole Kidman'}},

{'category': 'Space', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which planet is known for its Great Red Spot, a giant storm?', 'de': 'Welcher Planet ist fuer seinen Großen Roten Fleck bekannt, einen riesigen Sturm?', 'es': '¿Que planeta es conocido por su Gran Mancha Roja, una tormenta gigante?', 'nl': 'Welke planeet staat bekend om zijn Grote Rode Vlek, een gigantische storm?', 'sv': 'Vilken planet ar kand for sin Stora roda flacken, en jattestorm?', 'fr': "Quelle planete est connue pour sa Grande Tache rouge, une tempete geante ?", 'da': 'Hvilken planet er kendt for sin Store Røde Plet, en kæmpe storm?'},
 'opt1': {'en': 'Saturn', 'de': 'Saturn', 'es': 'Saturno', 'nl': 'Saturnus', 'sv': 'Saturnus', 'fr': 'Saturne', 'da': 'Saturn'},
 'opt2': {'en': 'Jupiter', 'de': 'Jupiter', 'es': 'Jupiter', 'nl': 'Jupiter', 'sv': 'Jupiter', 'fr': 'Jupiter', 'da': 'Jupiter'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which country's capital, Ulaanbaatar, is one of the coldest capital cities in the world?", 'de': 'Welches Land hat mit Ulaanbaatar eine der kaeltesten Hauptstaedte der Welt?', 'es': '¿Que pais tiene como capital a Ulan Bator, una de las capitales mas frias del mundo?', 'nl': 'Welk land heeft met Ulaanbaatar een van de koudste hoofdsteden ter wereld?', 'sv': 'Vilket lands huvudstad, Ulaanbaatar, ar en av varldens kallaste huvudstader?', 'fr': "Quel pays a pour capitale Oulan-Bator, l'une des capitales les plus froides du monde ?", 'da': 'Hvilket lands hovedstad, Ulaanbaatar, er en af verdens koldeste hovedstæder?'},
 'opt1': {'en': 'Mongolia', 'de': 'Mongolei', 'es': 'Mongolia', 'nl': 'Mongolie', 'sv': 'Mongoliet', 'fr': 'la Mongolie', 'da': 'Mongoliet'},
 'opt2': {'en': 'Kazakhstan', 'de': 'Kasachstan', 'es': 'Kazajistan', 'nl': 'Kazachstan', 'sv': 'Kazakstan', 'fr': 'le Kazakhstan', 'da': 'Kasakhstan'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Spanish designer is known for dramatic, sculptural haute couture gowns?', 'de': 'Welcher spanische Designer ist fuer dramatische, skulpturale Haute-Couture-Kleider bekannt?', 'es': '¿Que diseñador español es conocido por sus dramaticos y esculturales vestidos de alta costura?', 'nl': 'Welke Spaanse ontwerper staat bekend om dramatische, sculpturale haute couture jurken?', 'sv': 'Vilken spansk designer ar kand for dramatiska, skulpturala haute couture-klanningar?', 'fr': "Quel createur espagnol est connu pour ses robes de haute couture dramatiques et sculpturales ?", 'da': 'Hvilken spansk designer er kendt for dramatiske, skulpturelle haute couture-kjoler?'},
 'opt1': {'en': 'Cristobal Balenciaga', 'de': 'Cristobal Balenciaga', 'es': 'Cristobal Balenciaga', 'nl': 'Cristobal Balenciaga', 'sv': 'Cristobal Balenciaga', 'fr': 'Cristobal Balenciaga', 'da': 'Cristobal Balenciaga'},
 'opt2': {'en': 'Oscar de la Renta', 'de': 'Oscar de la Renta', 'es': 'Oscar de la Renta', 'nl': 'Oscar de la Renta', 'sv': 'Oscar de la Renta', 'fr': 'Oscar de la Renta', 'da': 'Oscar de la Renta'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which German philosopher wrote 'Critique of Pure Reason'?", 'de': "Welcher deutsche Philosoph schrieb 'Kritik der reinen Vernunft'?", 'es': "¿Que filosofo aleman escribio 'Critica de la razon pura'?", 'nl': "Welke Duitse filosoof schreef 'Kritiek van de zuivere rede'?", 'sv': "Vilken tysk filosof skrev 'Kritik av det rena fornuftet'?", 'fr': "Quel philosophe allemand a ecrit 'Critique de la raison pure' ?", 'da': "Hvilken tysk filosof skrev 'Kritik af den rene fornuft'?"},
 'opt1': {'en': 'Georg Hegel', 'de': 'Georg Hegel', 'es': 'Georg Hegel', 'nl': 'Georg Hegel', 'sv': 'Georg Hegel', 'fr': 'Georg Hegel', 'da': 'Georg Hegel'},
 'opt2': {'en': 'Immanuel Kant', 'de': 'Immanuel Kant', 'es': 'Immanuel Kant', 'nl': 'Immanuel Kant', 'sv': 'Immanuel Kant', 'fr': 'Emmanuel Kant', 'da': 'Immanuel Kant'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which quilting pattern uses a repeated pinwheel shape made from triangles?', 'de': 'Welches Quiltmuster verwendet eine wiederholte Windrad-Form aus Dreiecken?', 'es': '¿Que patron de acolchado usa una forma repetida de molinete hecha de triangulos?', 'nl': 'Welk quiltpatroon gebruikt een herhaalde windmolenvorm gemaakt van driehoeken?', 'sv': 'Vilket lapptacksmonster anvander en upprepad snurrform gjord av trianglar?', 'fr': "Quel motif de patchwork utilise une forme repetee de moulinet faite de triangles ?", 'da': 'Hvilket quiltmønster bruger en gentaget vindmøllefigur lavet af trekanter?'},
 'opt1': {'en': 'Pinwheel quilt', 'de': 'Windrad-Quilt', 'es': 'Edredon molinete', 'nl': 'Windmolenquilt', 'sv': 'Snurraquilt', 'fr': 'Patchwork moulinet', 'da': 'Vindmøllequilt'},
 'opt2': {'en': 'Ohio Star', 'de': 'Ohio-Stern', 'es': 'Estrella de Ohio', 'nl': 'Ohio Star', 'sv': 'Ohio Star', 'fr': 'Etoile de l\'Ohio', 'da': 'Ohio Star'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which flowering plant's bulb is famously associated with a 17th-century Dutch economic bubble?", 'de': 'Welche bluehende Pflanze mit einer Zwiebel wird beruehmt mit einer niederlaendischen Wirtschaftsblase des 17. Jahrhunderts assoziiert?', 'es': '¿Que planta con bulbo esta famosamente asociada con una burbuja economica holandesa del siglo XVII?', 'nl': 'Welke bloeiende plant met een bol wordt beroemd geassocieerd met een Nederlandse economische zeepbel uit de 17e eeuw?', 'sv': 'Vilken blommande vaxts lok forknippas beromt med en nederlandsk ekonomisk bubbla fran 1600-talet?', 'fr': "Quelle plante a fleurs dont le bulbe est celebrement associe a une bulle economique neerlandaise du 17e siecle ?", 'da': 'Hvilken blomstrende plantes løg er berømt forbundet med en hollandsk økonomisk boble fra det 17. århundrede?'},
 'opt1': {'en': 'Hyacinth', 'de': 'Hyazinthe', 'es': 'Jacinto', 'nl': 'Hyacint', 'sv': 'Hyacint', 'fr': 'Jacinthe', 'da': 'Hyacint'},
 'opt2': {'en': 'Tulip', 'de': 'Tulpe', 'es': 'Tulipan', 'nl': 'Tulp', 'sv': 'Tulpan', 'fr': 'Tulipe', 'da': 'Tulipan'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which architect designed the Eiffel Tower for the 1889 World's Fair?", 'de': "Welcher Architekt entwarf den Eiffelturm fuer die Weltausstellung 1889?", 'es': "¿Que arquitecto diseño la Torre Eiffel para la Exposicion Universal de 1889?", 'nl': "Welke architect ontwierp de Eiffeltoren voor de wereldtentoonstelling van 1889?", 'sv': "Vilken arkitekt designade Eiffeltornet for varldsutstallningen 1889?", 'fr': "Quel architecte a conçu la tour Eiffel pour l'Exposition universelle de 1889 ?", 'da': 'Hvilken arkitekt designede Eiffeltårnet til verdensudstillingen i 1889?'},
 'opt1': {'en': 'Gustave Eiffel', 'de': 'Gustave Eiffel', 'es': 'Gustave Eiffel', 'nl': 'Gustave Eiffel', 'sv': 'Gustave Eiffel', 'fr': 'Gustave Eiffel', 'da': 'Gustave Eiffel'},
 'opt2': {'en': 'Georges-Eugene Haussmann', 'de': 'Georges-Eugene Haussmann', 'es': 'Georges-Eugene Haussmann', 'nl': 'Georges-Eugene Haussmann', 'sv': 'Georges-Eugene Haussmann', 'fr': 'Georges-Eugene Haussmann', 'da': 'Georges-Eugene Haussmann'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'In Greek mythology, which goddess of the hunt is associated with the moon?', 'de': 'In der griechischen Mythologie, welche Jagdgoettin wird mit dem Mond assoziiert?', 'es': '¿En la mitologia griega, que diosa de la caza se asocia con la luna?', 'nl': 'In de Griekse mythologie, welke godin van de jacht wordt geassocieerd met de maan?', 'sv': 'I grekisk mytologi, vilken jaktgudinna forknippas med manen?', 'fr': "Dans la mythologie grecque, quelle deesse de la chasse est associee a la lune ?", 'da': 'I græsk mytologi, hvilken jagtgudinde forbindes med månen?'},
 'opt1': {'en': 'Selene', 'de': 'Selene', 'es': 'Selene', 'nl': 'Selene', 'sv': 'Selene', 'fr': 'Selene', 'da': 'Selene'},
 'opt2': {'en': 'Artemis', 'de': 'Artemis', 'es': 'Artemisa', 'nl': 'Artemis', 'sv': 'Artemis', 'fr': 'Artemis', 'da': 'Artemis'}},
]
