# -*- coding: utf-8 -*-
# Trivia Challenge batch #4. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which green herb is often used as a garnish and in tabbouleh salad?', 'de': 'Welches gruene Kraut wird oft als Garnitur und in Taboule-Salat verwendet?', 'es': '¿Que hierba verde se usa a menudo como guarnicion y en la ensalada tabule?', 'nl': 'Welk groen kruid wordt vaak gebruikt als garnering en in tabbouleh-salade?', 'sv': 'Vilken gron ort anvands ofta som garnering och i tabbouleh-sallad?', 'fr': "Quelle herbe verte est souvent utilisee en garniture et dans la salade taboule ?", 'da': 'Hvilken grøn urt bruges ofte som pynt og i tabbouleh-salat?'},
 'opt1': {'en': 'Parsley', 'de': 'Petersilie', 'es': 'Perejil', 'nl': 'Peterselie', 'sv': 'Persilja', 'fr': 'Persil', 'da': 'Persille'},
 'opt2': {'en': 'Cilantro', 'de': 'Koriander', 'es': 'Cilantro', 'nl': 'Koriander', 'sv': 'Koriander', 'fr': 'Coriandre', 'da': 'Koriander'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which country is home to the Amazon Rainforest's largest share?", 'de': 'Welches Land hat den groessten Anteil am Amazonas-Regenwald?', 'es': '¿Que pais tiene la mayor parte de la selva amazonica?', 'nl': 'Welk land heeft het grootste deel van het Amazoneregenwoud?', 'sv': 'Vilket land har den storsta andelen av Amazonas regnskog?', 'fr': "Quel pays possede la plus grande part de la foret amazonienne ?", 'da': 'Hvilket land har den største andel af Amazonas regnskov?'},
 'opt1': {'en': 'Brazil', 'de': 'Brasilien', 'es': 'Brasil', 'nl': 'Brazilie', 'sv': 'Brasilien', 'fr': 'le Bresil', 'da': 'Brasilien'},
 'opt2': {'en': 'Peru', 'de': 'Peru', 'es': 'Peru', 'nl': 'Peru', 'sv': 'Peru', 'fr': 'le Perou', 'da': 'Peru'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a shark terrorizing a beach town?', 'de': 'In welchem Filmklassiker terrorisiert ein Hai eine Strandstadt?', 'es': '¿Que pelicula clasica presenta a un tiburon aterrorizando un pueblo de playa?', 'nl': 'In welke klassieke film terroriseert een haai een strandstadje?', 'sv': 'I vilken klassisk film terroriserar en haj en strandstad?', 'fr': "Quel film classique met en scene un requin terrorisant une ville balneaire ?", 'da': 'Hvilken klassisk film har en haj, der terroriserer en strandby?'},
 'opt1': {'en': 'Jaws', 'de': 'Der weiße Hai', 'es': 'Tiburon', 'nl': 'Jaws', 'sv': 'Hajen', 'fr': 'Les Dents de la mer', 'da': 'Dødens gab'},
 'opt2': {'en': 'Deep Blue Sea', 'de': 'Deep Blue Sea', 'es': 'Deep Blue Sea', 'nl': 'Deep Blue Sea', 'sv': 'Deep Blue Sea', 'fr': 'Peur bleue', 'da': 'Deep Blue Sea'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which pet is known for purring and chasing mice?', 'de': 'Welches Haustier schnurrt und jagt Maeuse?', 'es': '¿Que mascota ronronea y persigue ratones?', 'nl': 'Welk huisdier spint en jaagt op muizen?', 'sv': 'Vilket sallskapsdjur spinner och jagar moss?', 'fr': "Quel animal de compagnie ronronne et chasse les souris ?", 'da': 'Hvilket kæledyr spinder og jagter mus?'},
 'opt1': {'en': 'Cat', 'de': 'Katze', 'es': 'Gato', 'nl': 'Kat', 'sv': 'Katt', 'fr': 'Chat', 'da': 'Kat'},
 'opt2': {'en': 'Dog', 'de': 'Hund', 'es': 'Perro', 'nl': 'Hond', 'sv': 'Hund', 'fr': 'Chien', 'da': 'Hund'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which frozen treat is made from cream, sugar, and flavorings?', 'de': 'Welche gefrorene Suessigkeit wird aus Sahne, Zucker und Aromen hergestellt?', 'es': '¿Que postre helado se hace con crema, azucar y aromas?', 'nl': 'Welk bevroren lekkernij wordt gemaakt van room, suiker en smaakstoffen?', 'sv': 'Vilken fryst godsak gors av gradde, socker och smaksattningar?', 'fr': "Quelle gourmandise glacee est faite de creme, de sucre et d'aromes ?", 'da': 'Hvilken frossen godbid er lavet af fløde, sukker og smagsstoffer?'},
 'opt1': {'en': 'Ice cream', 'de': 'Eiscreme', 'es': 'Helado', 'nl': 'IJs', 'sv': 'Glass', 'fr': 'Glace', 'da': 'Is'},
 'opt2': {'en': 'Sorbet', 'de': 'Sorbet', 'es': 'Sorbete', 'nl': 'Sorbet', 'sv': 'Sorbet', 'fr': 'Sorbet', 'da': 'Sorbet'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the Eiffel Tower and Louvre Museum?', 'de': 'In welchem Land befinden sich der Eiffelturm und der Louvre?', 'es': '¿En que pais estan la Torre Eiffel y el Museo del Louvre?', 'nl': 'In welk land bevinden zich de Eiffeltoren en het Louvre-museum?', 'sv': 'I vilket land ligger Eiffeltornet och Louvren?', 'fr': "Dans quel pays se trouvent la tour Eiffel et le musee du Louvre ?", 'da': 'Hvilket land huser Eiffeltårnet og Louvre-museet?'},
 'opt1': {'en': 'France', 'de': 'Frankreich', 'es': 'Francia', 'nl': 'Frankrijk', 'sv': 'Frankrike', 'fr': 'la France', 'da': 'Frankrig'},
 'opt2': {'en': 'Belgium', 'de': 'Belgien', 'es': 'Belgica', 'nl': 'Belgie', 'sv': 'Belgien', 'fr': 'la Belgique', 'da': 'Belgien'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which explorer led the first expedition to circumnavigate the globe, though he died before completing it?', 'de': 'Welcher Entdecker leitete die erste Weltumsegelung, starb aber vor deren Abschluss?', 'es': '¿Que explorador lidero la primera expedicion en circunnavegar el globo, aunque murio antes de completarla?', 'nl': 'Welke ontdekkingsreiziger leidde de eerste expeditie rond de wereld, hoewel hij stierf voordat deze voltooid was?', 'sv': 'Vilken upptacktsresande ledde den forsta expeditionen att segla jorden runt, fast han dog innan den avslutades?', 'fr': "Quel explorateur a dirige la premiere expedition faisant le tour du monde, bien qu'il soit mort avant son achevement ?", 'da': 'Hvilken opdagelsesrejsende ledte den første jordomsejling, selvom han døde, før den var fuldført?'},
 'opt1': {'en': 'Ferdinand Magellan', 'de': 'Ferdinand Magellan', 'es': 'Fernando de Magallanes', 'nl': 'Ferdinand Magellaan', 'sv': 'Ferdinand Magellan', 'fr': 'Fernand de Magellan', 'da': 'Ferdinand Magellan'},
 'opt2': {'en': 'Vasco da Gama', 'de': 'Vasco da Gama', 'es': 'Vasco da Gama', 'nl': 'Vasco da Gama', 'sv': 'Vasco da Gama', 'fr': 'Vasco de Gama', 'da': 'Vasco da Gama'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which art movement features bold brushstrokes and vivid color, led by Van Gogh and Monet?', 'de': 'Welche Kunstrichtung ist bekannt fuer kraeftige Pinselstriche und lebendige Farben, angefuehrt von Van Gogh und Monet?', 'es': '¿Que movimiento artistico presenta pinceladas audaces y colores vivos, liderado por Van Gogh y Monet?', 'nl': 'Welke kunststroming kenmerkt zich door gedurfde penseelstreken en levendige kleuren, geleid door Van Gogh en Monet?', 'sv': 'Vilken konstriktning kannetecknas av djarva penseldrag och levande farg, ledd av Van Gogh och Monet?', 'fr': "Quel mouvement artistique se caracterise par des coups de pinceau audacieux et des couleurs vives, mene par Van Gogh et Monet ?", 'da': 'Hvilken kunstretning har dristige penselstrøg og levende farve, ledet af Van Gogh og Monet?'},
 'opt1': {'en': 'Impressionism', 'de': 'Impressionismus', 'es': 'Impresionismo', 'nl': 'Impressionisme', 'sv': 'Impressionism', 'fr': 'Impressionnisme', 'da': 'Impressionisme'},
 'opt2': {'en': 'Realism', 'de': 'Realismus', 'es': 'Realismo', 'nl': 'Realisme', 'sv': 'Realism', 'fr': 'Realisme', 'da': 'Realisme'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which percussion instrument is played with two sticks and has a set of tuned metal bars?', 'de': 'Welches Perkussionsinstrument wird mit zwei Stoecken gespielt und hat gestimmte Metallstaebe?', 'es': '¿Que instrumento de percusion se toca con dos baquetas y tiene barras metalicas afinadas?', 'nl': 'Welk slaginstrument wordt met twee stokken bespeeld en heeft gestemde metalen staven?', 'sv': 'Vilket slaginstrument spelas med tva klubbor och har stamda metallstavar?', 'fr': "Quel instrument de percussion se joue avec deux baguettes et possede des lames metalliques accordees ?", 'da': 'Hvilket slagtøjsinstrument spilles med to stikker og har stemte metalstænger?'},
 'opt1': {'en': 'Xylophone', 'de': 'Xylophon', 'es': 'Xilofono', 'nl': 'Xylofoon', 'sv': 'Xylofon', 'fr': 'Xylophone', 'da': 'Xylofon'},
 'opt2': {'en': 'Marimba', 'de': 'Marimba', 'es': 'Marimba', 'nl': 'Marimba', 'sv': 'Marimba', 'fr': 'Marimba', 'da': 'Marimba'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which season is typically the coldest of the year?', 'de': 'Welche Jahreszeit ist typischerweise die kaelteste des Jahres?', 'es': '¿Que estacion es tipicamente la mas fria del año?', 'nl': 'Welk seizoen is doorgaans het koudste van het jaar?', 'sv': 'Vilken arstid ar vanligtvis den kallaste pa aret?', 'fr': "Quelle saison est typiquement la plus froide de l'annee ?", 'da': 'Hvilken årstid er typisk den koldeste på året?'},
 'opt1': {'en': 'Winter', 'de': 'Winter', 'es': 'Invierno', 'nl': 'Winter', 'sv': 'Vinter', 'fr': 'Hiver', 'da': 'Vinter'},
 'opt2': {'en': 'Autumn', 'de': 'Herbst', 'es': 'Otoño', 'nl': 'Herfst', 'sv': 'Host', 'fr': 'Automne', 'da': 'Efterår'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a governess who sings with the von Trapp children?', 'de': 'In welchem Filmklassiker singt eine Gouvernante mit den von-Trapp-Kindern?', 'es': '¿Que pelicula clasica presenta a una institutriz que canta con los niños von Trapp?', 'nl': 'In welke klassieke film zingt een gouvernante met de von Trapp-kinderen?', 'sv': 'I vilken klassisk film sjunger en guvernant med von Trapp-barnen?', 'fr': "Quel film classique met en scene une gouvernante qui chante avec les enfants von Trapp ?", 'da': 'Hvilken klassisk film har en guvernante, der synger med von Trapp-børnene?'},
 'opt1': {'en': 'The Sound of Music', 'de': 'Meine Lieder, meine Traeume', 'es': 'La Novicia Rebelde', 'nl': 'The Sound of Music', 'sv': 'Sound of Music', 'fr': 'La Melodie du bonheur', 'da': 'Sound of Music'},
 'opt2': {'en': 'Mary Poppins', 'de': 'Mary Poppins', 'es': 'Mary Poppins', 'nl': 'Mary Poppins', 'sv': 'Mary Poppins', 'fr': 'Mary Poppins', 'da': 'Mary Poppins'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which U.S. state is known as the 'Peach State'?", 'de': "Welcher US-Bundesstaat ist als 'Peach State' bekannt?", 'es': "¿Que estado de EE.UU. es conocido como el 'Estado del Melocoton'?", 'nl': "Welke Amerikaanse staat staat bekend als de 'Peach State'?", 'sv': "Vilken amerikansk delstat ar kand som 'Persikostaten'?", 'fr': "Quel État americain est connu comme le 'Peach State' (État de la peche) ?", 'da': "Hvilken amerikansk delstat er kendt som 'Peach State'?"},
 'opt1': {'en': 'Georgia', 'de': 'Georgia', 'es': 'Georgia', 'nl': 'Georgia', 'sv': 'Georgia', 'fr': 'la Georgie', 'da': 'Georgia'},
 'opt2': {'en': 'South Carolina', 'de': 'South Carolina', 'es': 'Carolina del Sur', 'nl': 'South Carolina', 'sv': 'South Carolina', 'fr': 'la Caroline du Sud', 'da': 'South Carolina'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which yellow vegetable is often grilled and used in Mexican dishes as 'elote'?", 'de': "Welches gelbe Gemuese wird oft gegrillt und in mexikanischen Gerichten als 'Elote' verwendet?", 'es': "¿Que verdura amarilla se asa a menudo y se usa en platos mexicanos como 'elote'?", 'nl': "Welke gele groente wordt vaak gegrild en gebruikt in Mexicaanse gerechten als 'elote'?", 'sv': "Vilken gul gronsak grillas ofta och anvands i mexikanska ratter som 'elote'?", 'fr': "Quel legume jaune est souvent grille et utilise dans les plats mexicains sous le nom d''elote' ?", 'da': "Hvilken gul grøntsag grilles ofte og bruges i mexicanske retter som 'elote'?"},
 'opt1': {'en': 'Corn', 'de': 'Mais', 'es': 'Maiz', 'nl': 'Mais', 'sv': 'Majs', 'fr': 'Le mais', 'da': 'Majs'},
 'opt2': {'en': 'Squash', 'de': 'Kuerbis', 'es': 'Calabaza', 'nl': 'Pompoen', 'sv': 'Squash', 'fr': 'La courge', 'da': 'Squash'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport uses a small white ball hit into holes with clubs across a large course?', 'de': 'Bei welcher Sportart wird ein kleiner weisser Ball mit Schlaegern in Loecher auf einem grossen Platz geschlagen?', 'es': '¿Que deporte usa una pequeña pelota blanca golpeada hacia hoyos con palos en un gran campo?', 'nl': 'Bij welke sport wordt een klein wit balletje met clubs in holes geslagen over een grote baan?', 'sv': 'Vilken sport anvander en liten vit boll som slas i hal med klubbor over en stor bana?', 'fr': "Quel sport utilise une petite balle blanche frappee dans des trous avec des clubs sur un grand parcours ?", 'da': 'Hvilken sport bruger en lille hvid bold, der slås i huller med køller på en stor bane?'},
 'opt1': {'en': 'Golf', 'de': 'Golf', 'es': 'Golf', 'nl': 'Golf', 'sv': 'Golf', 'fr': 'Golf', 'da': 'Golf'},
 'opt2': {'en': 'Croquet', 'de': 'Krocket', 'es': 'Croquet', 'nl': 'Croquet', 'sv': 'Krocket', 'fr': 'Croquet', 'da': 'Kroket'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which slow-moving reptile is known for its hard protective shell?', 'de': 'Welches sich langsam bewegende Reptil ist fuer seinen harten Schutzpanzer bekannt?', 'es': '¿Que reptil de movimiento lento es conocido por su duro caparazon protector?', 'nl': 'Welk langzaam bewegend reptiel staat bekend om zijn harde beschermende schild?', 'sv': 'Vilket langsamt reptil ar kant for sitt harda skyddande skal?', 'fr': "Quel reptile lent est connu pour sa carapace protectrice dure ?", 'da': 'Hvilket langsomt bevægende krybdyr er kendt for sit hårde beskyttende skjold?'},
 'opt1': {'en': 'Turtle', 'de': 'Schildkroete', 'es': 'Tortuga', 'nl': 'Schildpad', 'sv': 'Skoldpadda', 'fr': 'Tortue', 'da': 'Skildpadde'},
 'opt2': {'en': 'Lizard', 'de': 'Eidechse', 'es': 'Lagarto', 'nl': 'Hagedis', 'sv': 'Odla', 'fr': 'Lezard', 'da': 'Firben'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which holiday involves dressing up in green and celebrating Irish heritage in March?', 'de': 'Bei welchem Feiertag im Maerz kleidet man sich in gruen und feiert das irische Erbe?', 'es': '¿Que festividad implica vestirse de verde y celebrar la herencia irlandesa en marzo?', 'nl': 'Bij welke feestdag in maart kleedt men zich in groen en viert men de Ierse erfenis?', 'sv': 'Vilken hogtid i mars innebar att klä sig i gront och fira irlandskt arv?', 'fr': "Quelle fete implique de s'habiller en vert et de celebrer le patrimoine irlandais en mars ?", 'da': 'Hvilken helligdag i marts involverer at klæde sig i grønt og fejre irsk arv?'},
 'opt1': {'en': "St. Patrick's Day", 'de': 'St. Patrick\'s Day', 'es': 'Dia de San Patricio', 'nl': "St. Patrick's Day", 'sv': 'Sankt Patricks dag', 'fr': 'La Saint-Patrick', 'da': 'Sankt Patricks dag'},
 'opt2': {'en': 'Cinco de Mayo', 'de': 'Cinco de Mayo', 'es': 'Cinco de Mayo', 'nl': 'Cinco de Mayo', 'sv': 'Cinco de Mayo', 'fr': 'Cinco de Mayo', 'da': 'Cinco de Mayo'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the Sydney Opera House?', 'de': 'In welchem Land liegt das Sydney Opera House?', 'es': '¿En que pais esta la Opera de Sidney?', 'nl': 'In welk land ligt het Sydney Opera House?', 'sv': 'I vilket land ligger Sydneys operahus?', 'fr': "Dans quel pays se trouve l'Opera de Sydney ?", 'da': 'Hvilket land huser Sydney Opera House?'},
 'opt1': {'en': 'Australia', 'de': 'Australien', 'es': 'Australia', 'nl': 'Australie', 'sv': 'Australien', 'fr': "l'Australie", 'da': 'Australien'},
 'opt2': {'en': 'New Zealand', 'de': 'Neuseeland', 'es': 'Nueva Zelanda', 'nl': 'Nieuw-Zeeland', 'sv': 'Nya Zeeland', 'fr': 'la Nouvelle-Zelande', 'da': 'New Zealand'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which author wrote 'The Cat in the Hat'?", 'de': "Welcher Autor schrieb 'Der Kater mit Hut'?", 'es': "¿Que autor escribio 'El gato ensombrerado'?", 'nl': "Welke auteur schreef 'De Kat met de Hoed'?", 'sv': "Vilken forfattare skrev 'Katten med hatten'?", 'fr': "Quel auteur a ecrit 'Le Chat chapeaute' ?", 'da': "Hvilken forfatter skrev 'Katten med hatten'?"},
 'opt1': {'en': 'Dr. Seuss', 'de': 'Dr. Seuss', 'es': 'Dr. Seuss', 'nl': 'Dr. Seuss', 'sv': 'Dr. Seuss', 'fr': 'Dr Seuss', 'da': 'Dr. Seuss'},
 'opt2': {'en': 'Roald Dahl', 'de': 'Roald Dahl', 'es': 'Roald Dahl', 'nl': 'Roald Dahl', 'sv': 'Roald Dahl', 'fr': 'Roald Dahl', 'da': 'Roald Dahl'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic toy is a spinning disc thrown between players?', 'de': 'Welches klassische Spielzeug ist eine rotierende Scheibe, die zwischen Spielern geworfen wird?', 'es': '¿Que juguete clasico es un disco giratorio lanzado entre jugadores?', 'nl': 'Welk klassiek speelgoed is een draaiende schijf die tussen spelers wordt gegooid?', 'sv': 'Vilken klassisk leksak ar en snurrande skiva som kastas mellan spelare?', 'fr': "Quel jouet classique est un disque tournoyant lance entre les joueurs ?", 'da': 'Hvilket klassisk legetøj er en roterende skive, der kastes mellem spillere?'},
 'opt1': {'en': 'Frisbee', 'de': 'Frisbee', 'es': 'Frisbee', 'nl': 'Frisbee', 'sv': 'Frisbee', 'fr': 'Frisbee', 'da': 'Frisbee'},
 'opt2': {'en': 'Hula hoop', 'de': 'Hula-Hoop-Reifen', 'es': 'Hula hoop', 'nl': 'Hoelahoep', 'sv': 'Hoola-hoop', 'fr': 'Hula hoop', 'da': 'Hulahopring'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which spread is made from crushed olives and used in Mediterranean cooking?', 'de': 'Welcher Aufstrich wird aus zerdrueckten Oliven hergestellt und in der mediterranen Kueche verwendet?', 'es': '¿Que untable se hace con aceitunas trituradas y se usa en la cocina mediterranea?', 'nl': 'Welke smeersel wordt gemaakt van geplette olijven en gebruikt in de mediterrane keuken?', 'sv': 'Vilket pålagg gors av krossade oliver och anvands i medelhavsmat?', 'fr': "Quelle tartinade est faite d'olives ecrasees et utilisee dans la cuisine mediterraneenne ?", 'da': 'Hvilket pålæg er lavet af knuste oliven og bruges i middelhavsmad?'},
 'opt1': {'en': 'Tapenade', 'de': 'Tapenade', 'es': 'Tapenade', 'nl': 'Tapenade', 'sv': 'Tapenad', 'fr': 'Tapenade', 'da': 'Tapenade'},
 'opt2': {'en': 'Pesto', 'de': 'Pesto', 'es': 'Pesto', 'nl': 'Pesto', 'sv': 'Pesto', 'fr': 'Pesto', 'da': 'Pesto'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which wall divided a European city until it fell in 1989?', 'de': 'Welche Mauer teilte eine europaeische Stadt, bis sie 1989 fiel?', 'es': '¿Que muro dividio una ciudad europea hasta su caida en 1989?', 'nl': 'Welke muur verdeelde een Europese stad tot deze in 1989 viel?', 'sv': 'Vilken mur delade en europeisk stad tills den foll 1989?', 'fr': "Quel mur a divise une ville europeenne jusqu'a sa chute en 1989 ?", 'da': 'Hvilken mur delte en europæisk by, indtil den faldt i 1989?'},
 'opt1': {'en': 'The Berlin Wall', 'de': 'Die Berliner Mauer', 'es': 'El Muro de Berlin', 'nl': 'De Berlijnse Muur', 'sv': 'Berlinmuren', 'fr': 'Le mur de Berlin', 'da': 'Berlinmuren'},
 'opt2': {'en': "Hadrian's Wall", 'de': 'Hadrianswall', 'es': 'Muralla de Adriano', 'nl': 'De Muur van Hadrianus', 'sv': 'Hadrianus mur', 'fr': "Le mur d'Hadrien", 'da': 'Hadrians Mur'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which decade is disco's successor, new wave music, most associated with?", 'de': 'Mit welchem Jahrzehnt wird New-Wave-Musik, der Nachfolger von Disco, am meisten assoziiert?', 'es': '¿Con que decada se asocia mas la musica new wave, sucesora del disco?', 'nl': 'Met welk decennium wordt new wave-muziek, de opvolger van disco, het meest geassocieerd?', 'sv': 'Vilket artionde forknippas mest med new wave-musik, discons efterfoljare?', 'fr': "A quelle decennie associe-t-on le plus la new wave, successeure du disco ?", 'da': 'Hvilket årti forbindes mest med new wave-musik, discos efterfølger?'},
 'opt1': {'en': '1980s', 'de': '1980er', 'es': 'los años 80', 'nl': 'jaren 80', 'sv': '1980-talet', 'fr': 'les annees 1980', 'da': '1980’erne'},
 'opt2': {'en': '1970s', 'de': '1970er', 'es': 'los años 70', 'nl': 'jaren 70', 'sv': '1970-talet', 'fr': 'les annees 1970', 'da': '1970’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the Taj Mahal?', 'de': 'In welchem Land liegt der Taj Mahal?', 'es': '¿En que pais esta el Taj Mahal?', 'nl': 'In welk land ligt de Taj Mahal?', 'sv': 'I vilket land ligger Taj Mahal?', 'fr': "Dans quel pays se trouve le Taj Mahal ?", 'da': 'Hvilket land huser Taj Mahal?'},
 'opt1': {'en': 'India', 'de': 'Indien', 'es': 'India', 'nl': 'India', 'sv': 'Indien', 'fr': "l'Inde", 'da': 'Indien'},
 'opt2': {'en': 'Pakistan', 'de': 'Pakistan', 'es': 'Pakistan', 'nl': 'Pakistan', 'sv': 'Pakistan', 'fr': 'le Pakistan', 'da': 'Pakistan'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which farm animal produces milk used for cheese and yogurt?', 'de': 'Welches Nutztier produziert Milch fuer Kaese und Joghurt?', 'es': '¿Que animal de granja produce leche usada para queso y yogur?', 'nl': 'Welk boerderijdier produceert melk gebruikt voor kaas en yoghurt?', 'sv': 'Vilket lantbruksdjur producerar mjolk som anvands till ost och yoghurt?', 'fr': "Quel animal de ferme produit du lait utilise pour le fromage et le yaourt ?", 'da': 'Hvilket landbrugsdyr producerer mælk brugt til ost og yoghurt?'},
 'opt1': {'en': 'Cow', 'de': 'Kuh', 'es': 'Vaca', 'nl': 'Koe', 'sv': 'Ko', 'fr': 'Vache', 'da': 'Ko'},
 'opt2': {'en': 'Pig', 'de': 'Schwein', 'es': 'Cerdo', 'nl': 'Varken', 'sv': 'Gris', 'fr': 'Cochon', 'da': 'Gris'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which gemstone is the birthstone for those born in July?', 'de': 'Welcher Edelstein ist der Geburtsstein fuer im Juli Geborene?', 'es': '¿Que piedra preciosa es la piedra de nacimiento para los nacidos en julio?', 'nl': 'Welke edelsteen is de geboortesteen voor mensen geboren in juli?', 'sv': 'Vilken adelsten ar fodelsestenen for de fodda i juli?', 'fr': "Quelle pierre precieuse est la pierre de naissance de juillet ?", 'da': 'Hvilken ædelsten er fødselsstenen for dem, der er født i juli?'},
 'opt1': {'en': 'Ruby', 'de': 'Rubin', 'es': 'Rubi', 'nl': 'Robijn', 'sv': 'Rubin', 'fr': 'Rubis', 'da': 'Rubin'},
 'opt2': {'en': 'Sapphire', 'de': 'Saphir', 'es': 'Zafiro', 'nl': 'Saffier', 'sv': 'Safir', 'fr': 'Saphir', 'da': 'Safir'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic game involves guessing a hidden word one letter at a time?', 'de': 'Bei welchem klassischen Spiel raet man ein verstecktes Wort Buchstabe fuer Buchstabe?', 'es': '¿Que juego clasico consiste en adivinar una palabra oculta letra por letra?', 'nl': 'Bij welk klassiek spel raad je een verborgen woord letter voor letter?', 'sv': 'Vilket klassiskt spel innebar att gissa ett dolt ord en bokstav i taget?', 'fr': "Quel jeu classique consiste a deviner un mot cache lettre par lettre ?", 'da': 'Hvilket klassisk spil involverer at gætte et skjult ord et bogstav ad gangen?'},
 'opt1': {'en': 'Hangman', 'de': 'Galgenmaennchen', 'es': 'Ahorcado', 'nl': 'Galgje', 'sv': 'Hangman', 'fr': 'Le Pendu', 'da': 'Hængeren'},
 'opt2': {'en': 'Scrabble', 'de': 'Scrabble', 'es': 'Scrabble', 'nl': 'Scrabble', 'sv': 'Alfapet', 'fr': 'Scrabble', 'da': 'Scrabble'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which orange spice comes from a dried flower bud and is used in mulled cider?', 'de': 'Welches orangefarbene Gewuerz stammt von einer getrockneten Blueten knospe und wird in Gluehcider verwendet?', 'es': '¿Que especia naranja proviene de un capullo de flor seco y se usa en sidra especiada?', 'nl': 'Welke oranje specerij komt van een gedroogde bloemknop en wordt gebruikt in kruidencider?', 'sv': 'Vilken orange krydda kommer fran en torkad blomknopp och anvands i kryddad cider?', 'fr': "Quelle epice orange provient d'un bourgeon floral seche et est utilisee dans le cidre chaud epice ?", 'da': 'Hvilket orange krydderi kommer fra en tørret blomsterknop og bruges i krydret cider?'},
 'opt1': {'en': 'Clove', 'de': 'Gewuerznelke', 'es': 'Clavo', 'nl': 'Kruidnagel', 'sv': 'Kryddnejlika', 'fr': 'Le clou de girofle', 'da': 'Nellike'},
 'opt2': {'en': 'Allspice', 'de': 'Piment', 'es': 'Pimienta de Jamaica', 'nl': 'Piment', 'sv': 'Kryddpeppar', 'fr': 'Le quatre-epices', 'da': 'Allehånde'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which flower is traditionally associated with funerals and sympathy?', 'de': 'Welche Blume wird traditionell mit Beerdigungen und Anteilnahme assoziiert?', 'es': '¿Que flor se asocia tradicionalmente con funerales y condolencias?', 'nl': 'Welke bloem wordt traditioneel geassocieerd met begrafenissen en medeleven?', 'sv': 'Vilken blomma forknippas traditionellt med begravningar och medkansla?', 'fr': "Quelle fleur est traditionnellement associee aux funerailles et aux condoleances ?", 'da': 'Hvilken blomst forbindes traditionelt med begravelser og medfølelse?'},
 'opt1': {'en': 'Lily', 'de': 'Lilie', 'es': 'Lirio', 'nl': 'Lelie', 'sv': 'Lilja', 'fr': 'Lys', 'da': 'Lilje'},
 'opt2': {'en': 'Rose', 'de': 'Rose', 'es': 'Rosa', 'nl': 'Roos', 'sv': 'Ros', 'fr': 'Rose', 'da': 'Rose'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to Walt Disney World?', 'de': 'In welchem US-Bundesstaat liegt Walt Disney World?', 'es': '¿En que estado de EE.UU. esta Walt Disney World?', 'nl': 'In welke Amerikaanse staat ligt Walt Disney World?', 'sv': 'I vilken amerikansk delstat ligger Walt Disney World?', 'fr': "Dans quel État americain se trouve Walt Disney World ?", 'da': 'I hvilken amerikansk delstat ligger Walt Disney World?'},
 'opt1': {'en': 'Florida', 'de': 'Florida', 'es': 'Florida', 'nl': 'Florida', 'sv': 'Florida', 'fr': 'la Floride', 'da': 'Florida'},
 'opt2': {'en': 'California', 'de': 'Kalifornien', 'es': 'California', 'nl': 'Californie', 'sv': 'Kalifornien', 'fr': 'la Californie', 'da': 'Californien'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves pressing flowers between pages to preserve them?', 'de': 'Bei welchem Handwerk werden Blumen zwischen Seiten gepresst, um sie zu konservieren?', 'es': '¿Que manualidad implica prensar flores entre paginas para conservarlas?', 'nl': 'Bij welk handwerk worden bloemen tussen pagina\'s geperst om ze te bewaren?', 'sv': 'Vilket hantverk innebar att pressa blommor mellan sidor for att bevara dem?', 'fr': "Quel art manuel consiste a presser des fleurs entre des pages pour les conserver ?", 'da': 'Hvilket håndværk involverer at presse blomster mellem sider for at bevare dem?'},
 'opt1': {'en': 'Flower pressing', 'de': 'Blumenpressen', 'es': 'Prensado de flores', 'nl': 'Bloemen persen', 'sv': 'Blomsterpressning', 'fr': 'Le pressage de fleurs', 'da': 'Blomsterpresning'},
 'opt2': {'en': 'Flower arranging', 'de': 'Blumenbinden', 'es': 'Arreglo floral', 'nl': 'Bloemschikken', 'sv': 'Blomsterarrangering', 'fr': "L'art floral", 'da': 'Blomsterbinding'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves hitting a small ball against a wall with a racquet in an enclosed court?', 'de': 'Bei welcher Sportart wird ein kleiner Ball mit einem Schlaeger gegen eine Wand in einem geschlossenen Platz geschlagen?', 'es': '¿Que deporte consiste en golpear una pequeña pelota contra una pared con una raqueta en una cancha cerrada?', 'nl': 'Bij welke sport sla je een kleine bal met een racket tegen een muur in een gesloten baan?', 'sv': 'Vilken sport innebar att slå en liten boll mot en vagg med en racket i en inhagnad bana?', 'fr': "Quel sport consiste a frapper une petite balle contre un mur avec une raquette dans un court ferme ?", 'da': 'Hvilken sport involverer at slå en lille bold mod en væg med en ketcher i en lukket bane?'},
 'opt1': {'en': 'Racquetball', 'de': 'Racquetball', 'es': 'Racquetball', 'nl': 'Racquetball', 'sv': 'Racquetball', 'fr': 'Racquetball', 'da': 'Racquetball'},
 'opt2': {'en': 'Squash', 'de': 'Squash', 'es': 'Squash', 'nl': 'Squash', 'sv': 'Squash', 'fr': 'Squash', 'da': 'Squash'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which item is worn on the head to block sun, often with a wide brim?', 'de': 'Welches Kleidungsstueck wird auf dem Kopf getragen, um Sonne abzuhalten, oft mit breiter Krempe?', 'es': '¿Que prenda se lleva en la cabeza para bloquear el sol, a menudo con ala ancha?', 'nl': 'Welk item wordt op het hoofd gedragen om de zon te blokkeren, vaak met brede rand?', 'sv': 'Vilket plagg baers pa huvudet for att blockera solen, ofta med bred brätte?', 'fr': "Quel accessoire se porte sur la tete pour bloquer le soleil, souvent avec un large bord ?", 'da': 'Hvilken genstand bæres på hovedet for at blokere solen, ofte med bred skygge?'},
 'opt1': {'en': 'Sun hat', 'de': 'Sonnenhut', 'es': 'Sombrero de sol', 'nl': 'Zonnehoed', 'sv': 'Solhatt', 'fr': 'Chapeau de soleil', 'da': 'Solhat'},
 'opt2': {'en': 'Visor', 'de': 'Schirmmütze', 'es': 'Visera', 'nl': 'Zonneklep', 'sv': 'Skarm', 'fr': 'Visiere', 'da': 'Skygge'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic cartoon features a mouse named Mickey?', 'de': 'In welchem klassischen Cartoon kommt eine Maus namens Mickey vor?', 'es': '¿Que dibujo animado clasico presenta a un raton llamado Mickey?', 'nl': 'In welke klassieke tekenfilm komt een muis genaamd Mickey voor?', 'sv': 'I vilken klassisk tecknad film forekommer en mus vid namn Mickey?', 'fr': "Quel dessin anime classique met en scene une souris nommee Mickey ?", 'da': 'Hvilken klassisk tegnefilm har en mus ved navn Mickey?'},
 'opt1': {'en': 'Mickey Mouse', 'de': 'Micky Maus', 'es': 'Mickey Mouse', 'nl': 'Mickey Mouse', 'sv': 'Musse Pigg', 'fr': 'Mickey Mouse', 'da': 'Mickey Mouse'},
 'opt2': {'en': 'Tom and Jerry', 'de': 'Tom und Jerry', 'es': 'Tom y Jerry', 'nl': 'Tom en Jerry', 'sv': 'Tom och Jerry', 'fr': 'Tom et Jerry', 'da': 'Tom og Jerry'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to Table Mountain in Cape Town?', 'de': 'In welchem Land liegt der Tafelberg in Kapstadt?', 'es': '¿En que pais esta la Montaña de la Mesa en Ciudad del Cabo?', 'nl': 'In welk land ligt Tafelberg in Kaapstad?', 'sv': 'I vilket land ligger Taffelberget i Kapstaden?', 'fr': "Dans quel pays se trouve la montagne de la Table au Cap ?", 'da': 'Hvilket land huser Table Mountain i Cape Town?'},
 'opt1': {'en': 'South Africa', 'de': 'Suedafrika', 'es': 'Sudafrica', 'nl': 'Zuid-Afrika', 'sv': 'Sydafrika', 'fr': "l'Afrique du Sud", 'da': 'Sydafrika'},
 'opt2': {'en': 'Kenya', 'de': 'Kenia', 'es': 'Kenia', 'nl': 'Kenia', 'sv': 'Kenya', 'fr': 'le Kenya', 'da': 'Kenya'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady advocated for the arts and helped establish the National Endowment for the Arts era policies?', 'de': 'Welche First Lady setzte sich fuer die Kuenste ein und half, die Politik der National Endowment for the Arts Ära zu etablieren?', 'es': '¿Que primera dama abogo por las artes y ayudo a establecer las politicas de la era del National Endowment for the Arts?', 'nl': 'Welke First Lady zette zich in voor de kunsten en hielp bij het opzetten van beleid uit het National Endowment for the Arts-tijdperk?', 'sv': 'Vilken First Lady forespakade konst och hjalpte till att etablera National Endowment for the Arts-erans politik?', 'fr': "Quelle Première dame a defendu les arts et aide a etablir les politiques de l'ere du National Endowment for the Arts ?", 'da': 'Hvilken førstedame gik i bresche for kunsten og hjalp med at etablere National Endowment for the Arts-æraens politik?'},
 'opt1': {'en': 'Lady Bird Johnson', 'de': 'Lady Bird Johnson', 'es': 'Lady Bird Johnson', 'nl': 'Lady Bird Johnson', 'sv': 'Lady Bird Johnson', 'fr': 'Lady Bird Johnson', 'da': 'Lady Bird Johnson'},
 'opt2': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Spanish dish is a cold soup made from tomatoes, peppers, and cucumbers?', 'de': 'Welches spanische Gericht ist eine kalte Suppe aus Tomaten, Paprika und Gurken?', 'es': '¿Que plato español es una sopa fria hecha de tomates, pimientos y pepinos?', 'nl': 'Welk Spaans gerecht is een koude soep gemaakt van tomaten, paprika\'s en komkommers?', 'sv': 'Vilken spansk ratt ar en kall soppa gjord av tomater, paprika och gurka?', 'fr': "Quel plat espagnol est une soupe froide a base de tomates, poivrons et concombres ?", 'da': 'Hvilken spansk ret er en kold suppe lavet af tomater, peberfrugter og agurker?'},
 'opt1': {'en': 'Gazpacho', 'de': 'Gazpacho', 'es': 'Gazpacho', 'nl': 'Gazpacho', 'sv': 'Gazpacho', 'fr': 'Gaspacho', 'da': 'Gazpacho'},
 'opt2': {'en': 'Salmorejo', 'de': 'Salmorejo', 'es': 'Salmorejo', 'nl': 'Salmorejo', 'sv': 'Salmorejo', 'fr': 'Salmorejo', 'da': 'Salmorejo'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which plant structure absorbs water and nutrients from the soil?', 'de': 'Welche Pflanzenstruktur nimmt Wasser und Naehrstoffe aus dem Boden auf?', 'es': '¿Que estructura de la planta absorbe agua y nutrientes del suelo?', 'nl': 'Welke plantstructuur absorbeert water en voedingsstoffen uit de bodem?', 'sv': 'Vilken vaxtstruktur absorberar vatten och naringsamnen fran jorden?', 'fr': "Quelle structure vegetale absorbe l'eau et les nutriments du sol ?", 'da': 'Hvilken plantestruktur optager vand og næringsstoffer fra jorden?'},
 'opt1': {'en': 'Root', 'de': 'Wurzel', 'es': 'Raiz', 'nl': 'Wortel', 'sv': 'Rot', 'fr': 'Racine', 'da': 'Rod'},
 'opt2': {'en': 'Stem', 'de': 'Stengel', 'es': 'Tallo', 'nl': 'Stengel', 'sv': 'Stjalk', 'fr': 'Tige', 'da': 'Stængel'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which author wrote 'Wuthering Heights'?", 'de': "Welche Autorin schrieb 'Sturmhoehe'?", 'es': "¿Que autora escribio 'Cumbres Borrascosas'?", 'nl': "Welke auteur schreef 'Woeste Hoogten'?", 'sv': "Vilken forfattare skrev 'Svindlande hojder'?", 'fr': "Quelle autrice a ecrit 'Les Hauts de Hurlevent' ?", 'da': "Hvilken forfatter skrev 'Stormfulde højder'?"},
 'opt1': {'en': 'Emily Bronte', 'de': 'Emily Bronte', 'es': 'Emily Bronte', 'nl': 'Emily Bronte', 'sv': 'Emily Bronte', 'fr': 'Emily Bronte', 'da': 'Emily Bronte'},
 'opt2': {'en': 'Charlotte Bronte', 'de': 'Charlotte Bronte', 'es': 'Charlotte Bronte', 'nl': 'Charlotte Bronte', 'sv': 'Charlotte Bronte', 'fr': 'Charlotte Bronte', 'da': 'Charlotte Bronte'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Jo March in the 2019 film adaptation of 'Little Women'?", 'de': "Welche Schauspielerin spielte Jo March in der Filmadaption von 'Little Women' 2019?", 'es': "¿Que actriz interpreto a Jo March en la adaptacion cinematografica de 2019 de 'Mujercitas'?", 'nl': "Welke actrice speelde Jo March in de filmadaptatie van 'Little Women' uit 2019?", 'sv': "Vilken skadespelerska spelade Jo March i 2019 ars filmatisering av 'Unga kvinnor'?", 'fr': "Quelle actrice a joue Jo March dans l'adaptation cinematographique de 'Les Filles du docteur March' en 2019 ?", 'da': "Hvilken skuespiller spillede Jo March i 2019-filmatiseringen af 'Little Women'?"},
 'opt1': {'en': 'Saoirse Ronan', 'de': 'Saoirse Ronan', 'es': 'Saoirse Ronan', 'nl': 'Saoirse Ronan', 'sv': 'Saoirse Ronan', 'fr': 'Saoirse Ronan', 'da': 'Saoirse Ronan'},
 'opt2': {'en': 'Emma Watson', 'de': 'Emma Watson', 'es': 'Emma Watson', 'nl': 'Emma Watson', 'sv': 'Emma Watson', 'fr': 'Emma Watson', 'da': 'Emma Watson'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which singer is known as the 'King of Pop'?", 'de': "Welcher Saenger wird als 'King of Pop' bezeichnet?", 'es': "¿Que cantante es conocido como el 'Rey del Pop'?", 'nl': "Welke zanger staat bekend als de 'King of Pop'?", 'sv': "Vilken sangare kallas 'King of Pop'?", 'fr': "Quel chanteur est connu comme le 'Roi de la pop' ?", 'da': "Hvilken sanger er kendt som 'King of Pop'?"},
 'opt1': {'en': 'Michael Jackson', 'de': 'Michael Jackson', 'es': 'Michael Jackson', 'nl': 'Michael Jackson', 'sv': 'Michael Jackson', 'fr': 'Michael Jackson', 'da': 'Michael Jackson'},
 'opt2': {'en': 'Prince', 'de': 'Prince', 'es': 'Prince', 'nl': 'Prince', 'sv': 'Prince', 'fr': 'Prince', 'da': 'Prince'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which country is home to Angel Falls, the world's tallest waterfall?", 'de': 'In welchem Land liegt der Angel Falls, der hoechste Wasserfall der Welt?', 'es': '¿En que pais esta el Salto Angel, la catarata mas alta del mundo?', 'nl': 'In welk land ligt Angel Falls, de hoogste waterval ter wereld?', 'sv': 'I vilket land ligger Angel Falls, varldens hogsta vattenfall?', 'fr': "Dans quel pays se trouve le Salto Angel, la plus haute chute d'eau du monde ?", 'da': 'Hvilket land huser Angel Falls, verdens højeste vandfald?'},
 'opt1': {'en': 'Venezuela', 'de': 'Venezuela', 'es': 'Venezuela', 'nl': 'Venezuela', 'sv': 'Venezuela', 'fr': 'le Venezuela', 'da': 'Venezuela'},
 'opt2': {'en': 'Colombia', 'de': 'Kolumbien', 'es': 'Colombia', 'nl': 'Colombia', 'sv': 'Colombia', 'fr': 'la Colombie', 'da': 'Colombia'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large aquatic mammal is known for its tusks and whiskers, found in Arctic waters?', 'de': 'Welches grosse Wassersaeugetier ist fuer seine Stosszaehne und Barthaare bekannt und lebt in arktischen Gewaessern?', 'es': '¿Que gran mamifero acuatico es conocido por sus colmillos y bigotes, y se encuentra en aguas articas?', 'nl': 'Welk groot waterzoogdier staat bekend om zijn slagtanden en snorharen en leeft in Arctische wateren?', 'sv': 'Vilket stort vattenlevande daggdjur ar kant for sina betar och morrhar, och finns i arktiska vatten?', 'fr': "Quel grand mammifere aquatique est connu pour ses defenses et ses moustaches, trouve dans les eaux arctiques ?", 'da': 'Hvilket stort vandlevende pattedyr er kendt for sine stødtænder og overskæg, fundet i arktiske farvande?'},
 'opt1': {'en': 'Walrus', 'de': 'Walross', 'es': 'Morsa', 'nl': 'Walrus', 'sv': 'Valross', 'fr': 'Morse', 'da': 'Hvalros'},
 'opt2': {'en': 'Seal', 'de': 'Robbe', 'es': 'Foca', 'nl': 'Zeehond', 'sv': 'Sal', 'fr': 'Phoque', 'da': 'Sæl'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Italian bread is dipped in olive oil and often flavored with rosemary?', 'de': 'Welches italienische Brot wird in Olivenoel getunkt und oft mit Rosmarin gewuerzt?', 'es': '¿Que pan italiano se moja en aceite de oliva y a menudo se sazona con romero?', 'nl': 'Welk Italiaans brood wordt gedoopt in olijfolie en vaak op smaak gebracht met rozemarijn?', 'sv': 'Vilket italienskt brod dopps i olivolja och kryddas ofta med rosmarin?', 'fr': "Quel pain italien est trempe dans l'huile d'olive et souvent parfume au romarin ?", 'da': 'Hvilket italiensk brød dyppes i olivenolie og krydres ofte med rosmarin?'},
 'opt1': {'en': 'Focaccia', 'de': 'Focaccia', 'es': 'Focaccia', 'nl': 'Focaccia', 'sv': 'Focaccia', 'fr': 'Focaccia', 'da': 'Focaccia'},
 'opt2': {'en': 'Ciabatta', 'de': 'Ciabatta', 'es': 'Chapata', 'nl': 'Ciabatta', 'sv': 'Ciabatta', 'fr': 'Ciabatta', 'da': 'Ciabatta'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was the first to earn a graduate degree while serving as First Lady?', 'de': 'Welche First Lady erwarb als erste einen Hochschulabschluss waehrend ihrer Amtszeit als First Lady?', 'es': '¿Que primera dama fue la primera en obtener un titulo de posgrado mientras servia como primera dama?', 'nl': 'Welke First Lady was de eerste die een masterdiploma behaalde terwijl ze First Lady was?', 'sv': 'Vilken First Lady var forst med att ta en forskarutbildning medan hon tjanstgorde som First Lady?', 'fr': "Quelle Première dame a ete la premiere a obtenir un diplome d'etudes superieures pendant son mandat ?", 'da': 'Hvilken førstedame var den første til at tage en kandidatgrad, mens hun var førstedame?'},
 'opt1': {'en': 'Jill Biden', 'de': 'Jill Biden', 'es': 'Jill Biden', 'nl': 'Jill Biden', 'sv': 'Jill Biden', 'fr': 'Jill Biden', 'da': 'Jill Biden'},
 'opt2': {'en': 'Hillary Clinton', 'de': 'Hillary Clinton', 'es': 'Hillary Clinton', 'nl': 'Hillary Clinton', 'sv': 'Hillary Clinton', 'fr': 'Hillary Clinton', 'da': 'Hillary Clinton'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which artist is known for the painting 'American Gothic', depicting a farmer and his daughter?", 'de': "Welcher Kuenstler ist fuer das Gemaelde 'American Gothic' bekannt, das einen Farmer und seine Tochter zeigt?", 'es': "¿Que artista es conocido por la pintura 'American Gothic', que representa a un granjero y su hija?", 'nl': "Welke kunstenaar staat bekend om het schilderij 'American Gothic', dat een boer en zijn dochter afbeeldt?", 'sv': "Vilken konstnar ar kand for malningen 'American Gothic', som visar en bonde och hans dotter?", 'fr': "Quel artiste est connu pour le tableau 'American Gothic', representant un fermier et sa fille ?", 'da': "Hvilken kunstner er kendt for maleriet 'American Gothic', der viser en landmand og hans datter?"},
 'opt1': {'en': 'Grant Wood', 'de': 'Grant Wood', 'es': 'Grant Wood', 'nl': 'Grant Wood', 'sv': 'Grant Wood', 'fr': 'Grant Wood', 'da': 'Grant Wood'},
 'opt2': {'en': 'Edward Hopper', 'de': 'Edward Hopper', 'es': 'Edward Hopper', 'nl': 'Edward Hopper', 'sv': 'Edward Hopper', 'fr': 'Edward Hopper', 'da': 'Edward Hopper'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which flowering shrub is associated with the American South and blooms in shades of pink and white?', 'de': 'Welcher blueh ende Strauch wird mit dem amerikanischen Sueden assoziiert und blueht in Rosa- und Weisstoenen?', 'es': '¿Que arbusto en flor se asocia con el sur de EE.UU. y florece en tonos rosa y blanco?', 'nl': 'Welke bloeiende struik wordt geassocieerd met het Amerikaanse Zuiden en bloeit in tinten roze en wit?', 'sv': 'Vilken blommande buske forknippas med amerikanska sodern och blommar i rosa och vita nyanser?', 'fr': "Quel arbuste a fleurs est associe au Sud americain et fleurit dans des tons roses et blancs ?", 'da': 'Hvilken blomstrende busk forbindes med den amerikanske Syden og blomstrer i nuancer af pink og hvid?'},
 'opt1': {'en': 'Azalea', 'de': 'Azalee', 'es': 'Azalea', 'nl': 'Azalea', 'sv': 'Azalea', 'fr': 'Azalee', 'da': 'Azalea'},
 'opt2': {'en': 'Camellia', 'de': 'Kamelie', 'es': 'Camelia', 'nl': 'Camelia', 'sv': 'Kamelia', 'fr': 'Camelia', 'da': 'Kamelia'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which musical features the song 'Seasons of Love' and is set in New York's East Village?", 'de': "In welchem Musical kommt das Lied 'Seasons of Love' vor, das im East Village von New York spielt?", 'es': "¿Que musical presenta la cancion 'Seasons of Love' y se ambienta en el East Village de Nueva York?", 'nl': "Welke musical bevat het lied 'Seasons of Love' en speelt zich af in New Yorks East Village?", 'sv': "Vilken musikal innehaller latten 'Seasons of Love' och utspelar sig i New Yorks East Village?", 'fr': "Quelle comedie musicale comporte la chanson 'Seasons of Love' et se deroule dans l'East Village de New York ?", 'da': "Hvilken musical har sangen 'Seasons of Love' og foregår i New Yorks East Village?"},
 'opt1': {'en': 'Rent', 'de': 'Rent', 'es': 'Rent', 'nl': 'Rent', 'sv': 'Rent', 'fr': 'Rent', 'da': 'Rent'},
 'opt2': {'en': 'In the Heights', 'de': 'In the Heights', 'es': 'In the Heights', 'nl': 'In the Heights', 'sv': 'In the Heights', 'fr': 'In the Heights', 'da': 'In the Heights'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country has Warsaw as its capital?', 'de': 'Welches Land hat Warschau als Hauptstadt?', 'es': '¿Que pais tiene a Varsovia como su capital?', 'nl': 'Welk land heeft Warschau als hoofdstad?', 'sv': 'Vilket land har Warszawa som huvudstad?', 'fr': "Quel pays a Varsovie pour capitale ?", 'da': 'Hvilket land har Warszawa som hovedstad?'},
 'opt1': {'en': 'Poland', 'de': 'Polen', 'es': 'Polonia', 'nl': 'Polen', 'sv': 'Polen', 'fr': 'la Pologne', 'da': 'Polen'},
 'opt2': {'en': 'Czech Republic', 'de': 'Tschechien', 'es': 'Republica Checa', 'nl': 'Tsjechie', 'sv': 'Tjeckien', 'fr': 'la Republique tcheque', 'da': 'Tjekkiet'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which spice blend is central to Indian cuisine and includes cumin, coriander, and turmeric?', 'de': 'Welche Gewuerzmischung ist zentral fuer die indische Kueche und enthaelt Kreuzkuemmel, Koriander und Kurkuma?', 'es': '¿Que mezcla de especias es central en la cocina india e incluye comino, cilantro y curcuma?', 'nl': 'Welke kruidenmix is centraal in de Indiase keuken en bevat komijn, koriander en kurkuma?', 'sv': 'Vilken kryddblandning ar central i indisk mat och innehaller spiskummin, koriander och gurkmeja?', 'fr': "Quel melange d'epices est central dans la cuisine indienne et comprend cumin, coriandre et curcuma ?", 'da': 'Hvilken krydderiblanding er central i indisk mad og inkluderer spidskommen, koriander og gurkemeje?'},
 'opt1': {'en': 'Garam masala', 'de': 'Garam Masala', 'es': 'Garam masala', 'nl': 'Garam masala', 'sv': 'Garam masala', 'fr': 'Garam masala', 'da': 'Garam masala'},
 'opt2': {'en': 'Curry powder', 'de': 'Currypulver', 'es': 'Curry en polvo', 'nl': 'Kerriepoeder', 'sv': 'Curryblandning', 'fr': 'Poudre de curry', 'da': 'Karrypulver'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which needlework craft involves creating lace by knotting and looping thread with a small shuttle?', 'de': 'Welches Handarbeitshandwerk erzeugt Spitze durch Knoten und Schlingen von Faden mit einem kleinen Schiffchen?', 'es': '¿Que arte de aguja crea encaje anudando y enlazando hilo con una pequeña lanzadera?', 'nl': 'Welk naaiwerk creëert kant door draad te knopen en te lussen met een klein schuitje?', 'sv': 'Vilket handarbete skapar spets genom att knyta och sla trad med en liten spole?', 'fr': "Quel art de l'aiguille cree de la dentelle en nouant et bouclant du fil avec une petite navette ?", 'da': 'Hvilket håndarbejde skaber knipling ved at knytte og sløjfe tråd med en lille spole?'},
 'opt1': {'en': 'Tatting', 'de': 'Schiffchenspitze', 'es': 'Frivolite', 'nl': 'Frivolite', 'sv': 'Frivolitet', 'fr': 'La frivolite', 'da': 'Frivolitet'},
 'opt2': {'en': 'Crochet', 'de': 'Haekeln', 'es': 'Ganchillo', 'nl': 'Haken', 'sv': 'Virkning', 'fr': 'Le crochet', 'da': 'Hækling'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which organ produces bile to aid digestion?', 'de': 'Welches Organ produziert Galle zur Unterstuetzung der Verdauung?', 'es': '¿Que organo produce bilis para ayudar a la digestion?', 'nl': 'Welk orgaan produceert gal om de spijsvertering te ondersteunen?', 'sv': 'Vilket organ producerar galla for att hjalpa matsmaltningen?', 'fr': "Quel organe produit la bile pour aider la digestion ?", 'da': 'Hvilket organ producerer galde for at hjælpe fordøjelsen?'},
 'opt1': {'en': 'Liver', 'de': 'Leber', 'es': 'Higado', 'nl': 'Lever', 'sv': 'Lever', 'fr': 'Foie', 'da': 'Lever'},
 'opt2': {'en': 'Pancreas', 'de': 'Bauchspeicheldruese', 'es': 'Pancreas', 'nl': 'Alvleesklier', 'sv': 'Bukspottkortel', 'fr': 'Pancreas', 'da': 'Bugspytkirtel'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'In swimming, which stroke involves both arms moving simultaneously over the head in a circular motion?', 'de': 'Bei welchem Schwimmstil bewegen sich beide Arme gleichzeitig ueber dem Kopf in einer kreisenden Bewegung?', 'es': '¿En natacion, que estilo implica que ambos brazos se muevan simultaneamente sobre la cabeza en un movimiento circular?', 'nl': 'Bij welke zwemslag bewegen beide armen tegelijk over het hoofd in een cirkelvormige beweging?', 'sv': 'Vilken simstil innebar att bada armarna ror sig samtidigt over huvudet i en cirkelrorelse?', 'fr': "En natation, quelle nage implique que les deux bras se deplacent simultanement au-dessus de la tete dans un mouvement circulaire ?", 'da': 'Hvilken svømmestil involverer, at begge arme bevæger sig samtidigt over hovedet i en cirkulær bevægelse?'},
 'opt1': {'en': 'Butterfly', 'de': 'Schmetterling', 'es': 'Mariposa', 'nl': 'Vlinderslag', 'sv': 'Fjarilsim', 'fr': 'Papillon', 'da': 'Butterfly'},
 'opt2': {'en': 'Freestyle', 'de': 'Freistil', 'es': 'Estilo libre', 'nl': 'Vrije slag', 'sv': 'Frisim', 'fr': 'Nage libre', 'da': 'Fri stil'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Spanish city is known for the Sagrada Familia basilica?', 'de': 'Welche spanische Stadt ist fuer die Basilika Sagrada Familia bekannt?', 'es': '¿Que ciudad española es conocida por la basilica de la Sagrada Familia?', 'nl': 'Welke Spaanse stad staat bekend om de basiliek Sagrada Familia?', 'sv': 'Vilken spansk stad ar kand for basilikan Sagrada Familia?', 'fr': "Quelle ville espagnole est connue pour la basilique de la Sagrada Familia ?", 'da': 'Hvilken spansk by er kendt for basilikaen Sagrada Familia?'},
 'opt1': {'en': 'Barcelona', 'de': 'Barcelona', 'es': 'Barcelona', 'nl': 'Barcelona', 'sv': 'Barcelona', 'fr': 'Barcelone', 'da': 'Barcelona'},
 'opt2': {'en': 'Madrid', 'de': 'Madrid', 'es': 'Madrid', 'nl': 'Madrid', 'sv': 'Madrid', 'fr': 'Madrid', 'da': 'Madrid'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Katniss Everdeen in 'The Hunger Games'?", 'de': "Welche Schauspielerin spielte Katniss Everdeen in 'Die Tribute von Panem'?", 'es': "¿Que actriz interpreto a Katniss Everdeen en 'Los Juegos del Hambre'?", 'nl': "Welke actrice speelde Katniss Everdeen in 'The Hunger Games'?", 'sv': "Vilken skadespelerska spelade Katniss Everdeen i 'Hunger Games'?", 'fr': "Quelle actrice a joue Katniss Everdeen dans 'Hunger Games' ?", 'da': "Hvilken skuespiller spillede Katniss Everdeen i 'Hunger Games'?"},
 'opt1': {'en': 'Jennifer Lawrence', 'de': 'Jennifer Lawrence', 'es': 'Jennifer Lawrence', 'nl': 'Jennifer Lawrence', 'sv': 'Jennifer Lawrence', 'fr': 'Jennifer Lawrence', 'da': 'Jennifer Lawrence'},
 'opt2': {'en': 'Shailene Woodley', 'de': 'Shailene Woodley', 'es': 'Shailene Woodley', 'nl': 'Shailene Woodley', 'sv': 'Shailene Woodley', 'fr': 'Shailene Woodley', 'da': 'Shailene Woodley'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Charles Dickens novel features the orphan Pip and the mysterious Miss Havisham?', 'de': 'In welchem Roman von Charles Dickens kommen der Waise Pip und die mysterioese Miss Havisham vor?', 'es': '¿Que novela de Charles Dickens presenta al huerfano Pip y a la misteriosa Miss Havisham?', 'nl': 'In welke roman van Charles Dickens komen de wees Pip en de mysterieuze Miss Havisham voor?', 'sv': 'I vilken roman av Charles Dickens forekommer den foraldralosa Pip och den mystiska Miss Havisham?', 'fr': "Quel roman de Charles Dickens met en scene l'orphelin Pip et la mysterieuse Miss Havisham ?", 'da': 'Hvilken roman af Charles Dickens har den forældreløse Pip og den mystiske frøken Havisham?'},
 'opt1': {'en': 'Great Expectations', 'de': 'Große Erwartungen', 'es': 'Grandes esperanzas', 'nl': 'Great Expectations', 'sv': 'Lysande utsikter', 'fr': 'Les Grandes Esperances', 'da': 'Store forventninger'},
 'opt2': {'en': 'Oliver Twist', 'de': 'Oliver Twist', 'es': 'Oliver Twist', 'nl': 'Oliver Twist', 'sv': 'Oliver Twist', 'fr': 'Oliver Twist', 'da': 'Oliver Twist'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which 1960s band is known for the album 'Sgt. Pepper's Lonely Hearts Club Band'?", 'de': "Welche Band der 1960er ist fuer das Album 'Sgt. Pepper's Lonely Hearts Club Band' bekannt?", 'es': "¿Que banda de los 60 es conocida por el album 'Sgt. Pepper's Lonely Hearts Club Band'?", 'nl': "Welke band uit de jaren 60 staat bekend om het album 'Sgt. Pepper's Lonely Hearts Club Band'?", 'sv': "Vilket 1960-talsband ar kant for albumet 'Sgt. Pepper's Lonely Hearts Club Band'?", 'fr': "Quel groupe des annees 1960 est connu pour l'album 'Sgt. Pepper's Lonely Hearts Club Band' ?", 'da': "Hvilket 1960’er-band er kendt for albummet 'Sgt. Pepper's Lonely Hearts Club Band'?"},
 'opt1': {'en': 'The Beatles', 'de': 'The Beatles', 'es': 'The Beatles', 'nl': 'The Beatles', 'sv': 'The Beatles', 'fr': 'The Beatles', 'da': 'The Beatles'},
 'opt2': {'en': 'The Rolling Stones', 'de': 'The Rolling Stones', 'es': 'The Rolling Stones', 'nl': 'The Rolling Stones', 'sv': 'The Rolling Stones', 'fr': 'The Rolling Stones', 'da': 'The Rolling Stones'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a casserole made with green beans and fried onions, popular at Thanksgiving?', 'de': 'Welches suddstaatliche Gericht ist ein Auflauf aus gruenen Bohnen und gebratenen Zwiebeln, beliebt zu Thanksgiving?', 'es': '¿Que plato sureño de EE.UU. es una cacerola hecha con judias verdes y cebollas fritas, popular en Accion de Gracias?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een ovenschotel gemaakt met sperziebonen en gefrituurde uien, populair met Thanksgiving?', 'sv': 'Vilken sydstatlig ratt ar en gryta gjord med gronsaker och friterad lok, populart pa Thanksgiving?', 'fr': "Quel plat du Sud des États-Unis est une gratin de haricots verts et d'oignons frits, populaire a Thanksgiving ?", 'da': 'Hvilken sydstatsret er en gryderet lavet med grønne bønner og stegte løg, populær til Thanksgiving?'},
 'opt1': {'en': 'Green bean casserole', 'de': 'Gruene-Bohnen-Auflauf', 'es': 'Cacerola de judias verdes', 'nl': 'Sperziebonenschotel', 'sv': 'Grona bonor-gryta', 'fr': 'Le gratin de haricots verts', 'da': 'Grøn bønne-gryderet'},
 'opt2': {'en': 'Sweet potato casserole', 'de': 'Suesskartoffel-Auflauf', 'es': 'Cacerola de camote', 'nl': 'Zoete aardappelschotel', 'sv': 'Sotpotatisgryta', 'fr': 'Le gratin de patates douces', 'da': 'Sødkartoffel-gryderet'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which queen ruled ancient Egypt and famously allied with both Julius Caesar and Mark Antony?', 'de': 'Welche Koenigin regierte das antike Ägypten und verbuendete sich beruehmt sowohl mit Julius Caesar als auch mit Marcus Antonius?', 'es': '¿Que reina goberno el antiguo Egipto y se alio famosamente tanto con Julio Cesar como con Marco Antonio?', 'nl': 'Welke koningin regeerde over het oude Egypte en werd beroemd bondgenoot van zowel Julius Caesar als Marcus Antonius?', 'sv': 'Vilken drottning styrde det antika Egypten och allierade sig beromt med bade Julius Caesar och Marcus Antonius?', 'fr': "Quelle reine a regne sur l'Egypte antique et s'est celebrement alliee a la fois avec Jules Cesar et Marc Antoine ?", 'da': 'Hvilken dronning regerede det gamle Egypten og alliancerede sig berømt med både Julius Cæsar og Marcus Antonius?'},
 'opt1': {'en': 'Cleopatra', 'de': 'Kleopatra', 'es': 'Cleopatra', 'nl': 'Cleopatra', 'sv': 'Kleopatra', 'fr': 'Cleopatre', 'da': 'Kleopatra'},
 'opt2': {'en': 'Nefertiti', 'de': 'Nofretete', 'es': 'Nefertiti', 'nl': 'Nefertiti', 'sv': 'Nefertiti', 'fr': 'Nefertiti', 'da': 'Nefertiti'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large bird of prey has excellent eyesight and can dive at speeds over 200 mph while hunting?', 'de': 'Welcher grosse Greifvogel hat ausgezeichnete Sehkraft und kann beim Jagen mit ueber 320 km/h herabstuerzen?', 'es': '¿Que gran ave rapaz tiene una vista excelente y puede zambullirse a mas de 320 km/h al cazar?', 'nl': 'Welke grote roofvogel heeft uitstekend zicht en kan tijdens de jacht duiken met snelheden van meer dan 320 km/u?', 'sv': 'Vilken stor rovfagel har utmarkt syn och kan dyka i hastigheter over 320 km/h under jakt?', 'fr': "Quel grand rapace a une excellente vue et peut plonger a plus de 320 km/h en chassant ?", 'da': 'Hvilken stor rovfugl har fremragende syn og kan dykke med hastigheder over 320 km/t under jagt?'},
 'opt1': {'en': 'Peregrine falcon', 'de': 'Wanderfalke', 'es': 'Halcon peregrino', 'nl': 'Slechtvalk', 'sv': 'Pilgrimsfalk', 'fr': 'Faucon pelerin', 'da': 'Vandrefalk'},
 'opt2': {'en': 'Golden eagle', 'de': 'Steinadler', 'es': 'Aguila real', 'nl': 'Steenarend', 'sv': 'Kungsorn', 'fr': 'Aigle royal', 'da': 'Kongeørn'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient ruins of Chichen Itza?', 'de': 'In welchem Land liegen die antiken Ruinen von Chichen Itza?', 'es': '¿En que pais estan las ruinas antiguas de Chichen Itza?', 'nl': 'In welk land liggen de oude ruines van Chichen Itza?', 'sv': 'I vilket land ligger de antika ruinerna av Chichen Itza?', 'fr': "Dans quel pays se trouvent les ruines antiques de Chichen Itza ?", 'da': 'Hvilket land huser de gamle ruiner af Chichen Itza?'},
 'opt1': {'en': 'Mexico', 'de': 'Mexiko', 'es': 'Mexico', 'nl': 'Mexico', 'sv': 'Mexiko', 'fr': 'le Mexique', 'da': 'Mexico'},
 'opt2': {'en': 'Guatemala', 'de': 'Guatemala', 'es': 'Guatemala', 'nl': 'Guatemala', 'sv': 'Guatemala', 'fr': 'le Guatemala', 'da': 'Guatemala'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of sunglasses, popular in the 1950s-60s, features upswept, pointed frame edges?', 'de': 'Welche Art von Sonnenbrille, beliebt in den 1950er-60er Jahren, hat nach oben geschwungene, spitze Fassungsraender?', 'es': '¿Que tipo de gafas de sol, populares en los años 50-60, tiene bordes de montura puntiagudos y hacia arriba?', 'nl': 'Welk type zonnebril, populair in de jaren 50-60, heeft opwaarts gezwaaide, puntige montuurranden?', 'sv': 'Vilken typ av solglasogon, populara pa 1950-60-talen, har uppatsvangda, spetsiga bagkanter?', 'fr': "Quel type de lunettes de soleil, populaire dans les annees 1950-60, presente des bords de monture releves et pointus ?", 'da': 'Hvilken type solbriller, populære i 1950’erne-60’erne, har opadvendte, spidse stelkanter?'},
 'opt1': {'en': 'Cat-eye glasses', 'de': 'Cat-Eye-Brille', 'es': 'Gafas de ojo de gato', 'nl': 'Cat-eye bril', 'sv': 'Cateye-glasogon', 'fr': "Lunettes oeil de chat", 'da': 'Cateye-briller'},
 'opt2': {'en': 'Aviators', 'de': 'Pilotenbrille', 'es': 'Aviadores', 'nl': 'Aviator-zonnebril', 'sv': 'Pilotglasogon', 'fr': 'Aviateurs', 'da': 'Pilotbriller'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which Bridgerton spin-off series focuses on Queen Charlotte's younger years?", 'de': "Welche Bridgerton-Spin-off-Serie konzentriert sich auf die juengeren Jahre von Koenigin Charlotte?", 'es': "¿Que serie derivada de Bridgerton se centra en los años jovenes de la Reina Charlotte?", 'nl': "Welke Bridgerton-spin-offserie richt zich op de jongere jaren van Koningin Charlotte?", 'sv': "Vilken Bridgerton-spinoff-serie fokuserar pa drottning Charlottes yngre ar?", 'fr': "Quelle serie derivee de Bridgerton se concentre sur les jeunes annees de la reine Charlotte ?", 'da': 'Hvilken Bridgerton-spinoff-serie fokuserer på Dronning Charlottes yngre år?'},
 'opt1': {'en': 'Queen Charlotte: A Bridgerton Story', 'de': 'Queen Charlotte: A Bridgerton Story', 'es': 'La Reina Charlotte: Una Historia Bridgerton', 'nl': 'Queen Charlotte: A Bridgerton Story', 'sv': 'Queen Charlotte: A Bridgerton Story', 'fr': 'Queen Charlotte : Une histoire Bridgerton', 'da': 'Queen Charlotte: A Bridgerton Story'},
 'opt2': {'en': 'The Crown', 'de': 'The Crown', 'es': 'The Crown', 'nl': 'The Crown', 'sv': 'The Crown', 'fr': 'The Crown', 'da': 'The Crown'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which tree produces the acorns eaten by squirrels and is a symbol of strength?', 'de': 'Welcher Baum produziert die Eicheln, die von Eichhoernchen gefressen werden, und ist ein Symbol der Staerke?', 'es': '¿Que arbol produce las bellotas que comen las ardillas y es un simbolo de fortaleza?', 'nl': 'Welke boom produceert de eikels die door eekhoorns worden gegeten en is een symbool van kracht?', 'sv': 'Vilket trad producerar ekollon som ats av ekorrar och ar en styrkesymbol?', 'fr': "Quel arbre produit les glands manges par les ecureuils et est un symbole de force ?", 'da': 'Hvilket træ producerer agern, der spises af egern, og er et symbol på styrke?'},
 'opt1': {'en': 'Oak', 'de': 'Eiche', 'es': 'Roble', 'nl': 'Eik', 'sv': 'Ek', 'fr': 'Chene', 'da': 'Eg'},
 'opt2': {'en': 'Birch', 'de': 'Birke', 'es': 'Abedul', 'nl': 'Berk', 'sv': 'Bjork', 'fr': 'Bouleau', 'da': 'Birk'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Chinese dumpling is steamed and often filled with pork or shrimp?', 'de': 'Welcher chinesische Kloss wird gedaempft und oft mit Schweinefleisch oder Garnelen gefuellt?', 'es': '¿Que empanadilla china se cocina al vapor y a menudo se rellena con cerdo o camarones?', 'nl': 'Welke Chinese knoedel wordt gestoomd en vaak gevuld met varkensvlees of garnalen?', 'sv': 'Vilken kinesisk dumpling angkoks och fylls ofta med flask eller raga?', 'fr': "Quel dumpling chinois est cuit a la vapeur et souvent fourre au porc ou aux crevettes ?", 'da': 'Hvilken kinesisk dumpling dampes og fyldes ofte med svinekød eller rejer?'},
 'opt1': {'en': 'Dim sum dumpling', 'de': 'Dim-Sum-Kloss', 'es': 'Dim sum', 'nl': 'Dim sum-knoedel', 'sv': 'Dim sum-dumpling', 'fr': 'Dumpling dim sum', 'da': 'Dim sum-dumpling'},
 'opt2': {'en': 'Wonton', 'de': 'Wan Tan', 'es': 'Wonton', 'nl': 'Wonton', 'sv': 'Wonton', 'fr': 'Wonton', 'da': 'Wonton'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which sense allows humans to detect sound waves?', 'de': 'Welcher Sinn ermoeglicht es Menschen, Schallwellen wahrzunehmen?', 'es': '¿Que sentido permite a los humanos detectar ondas sonoras?', 'nl': 'Welk zintuig stelt mensen in staat geluidsgolven waar te nemen?', 'sv': 'Vilket sinne later manniskor upptacka ljudvagor?', 'fr': "Quel sens permet aux humains de detecter les ondes sonores ?", 'da': 'Hvilken sans lader mennesker opfange lydbølger?'},
 'opt1': {'en': 'Hearing', 'de': 'Hoeren', 'es': 'Audicion', 'nl': 'Gehoor', 'sv': 'Horsel', 'fr': "L'ouïe", 'da': 'Hørelse'},
 'opt2': {'en': 'Touch', 'de': 'Tasten', 'es': 'Tacto', 'nl': 'Tastzin', 'sv': 'Kansel', 'fr': 'Le toucher', 'da': 'Følesans'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Ephesus?', 'de': 'In welchem Land liegt die antike Stadt Ephesus?', 'es': '¿En que pais esta la antigua ciudad de Efeso?', 'nl': 'In welk land ligt de oude stad Efeze?', 'sv': 'I vilket land ligger den antika staden Efesos?', 'fr': "Dans quel pays se trouve l'ancienne cite d'Ephese ?", 'da': 'Hvilket land huser den antikke by Efesos?'},
 'opt1': {'en': 'Turkey', 'de': 'Tuerkei', 'es': 'Turquia', 'nl': 'Turkije', 'sv': 'Turkiet', 'fr': 'la Turquie', 'da': 'Tyrkiet'},
 'opt2': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was the first to have served in the U.S. Senate?', 'de': 'Welche First Lady war die erste, die im US-Senat gedient hatte?', 'es': '¿Que primera dama fue la primera en haber servido en el Senado de EE.UU.?', 'nl': 'Welke First Lady was de eerste die in de Amerikaanse Senaat had gediend?', 'sv': 'Vilken First Lady var forst med att ha tjanstgjort i USA:s senat?', 'fr': "Quelle Première dame a ete la premiere a avoir siege au Senat americain ?", 'da': 'Hvilken førstedame var den første til at have siddet i det amerikanske senat?'},
 'opt1': {'en': 'Hillary Clinton', 'de': 'Hillary Clinton', 'es': 'Hillary Clinton', 'nl': 'Hillary Clinton', 'sv': 'Hillary Clinton', 'fr': 'Hillary Clinton', 'da': 'Hillary Clinton'},
 'opt2': {'en': 'Elizabeth Dole', 'de': 'Elizabeth Dole', 'es': 'Elizabeth Dole', 'nl': 'Elizabeth Dole', 'sv': 'Elizabeth Dole', 'fr': 'Elizabeth Dole', 'da': 'Elizabeth Dole'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which composer wrote 'The Planets' orchestral suite?", 'de': "Welcher Komponist schrieb die Orchestersuite 'Die Planeten'?", 'es': "¿Que compositor escribio la suite orquestal 'Los planetas'?", 'nl': "Welke componist schreef de orkestsuite 'De Planeten'?", 'sv': "Vilken kompositor skrev orkestersviten 'Planeterna'?", 'fr': "Quel compositeur a ecrit la suite orchestrale 'Les Planetes' ?", 'da': "Hvilken komponist skrev orkestersuiten 'Planeterne'?"},
 'opt1': {'en': 'Gustav Holst', 'de': 'Gustav Holst', 'es': 'Gustav Holst', 'nl': 'Gustav Holst', 'sv': 'Gustav Holst', 'fr': 'Gustav Holst', 'da': 'Gustav Holst'},
 'opt2': {'en': 'Edward Elgar', 'de': 'Edward Elgar', 'es': 'Edward Elgar', 'nl': 'Edward Elgar', 'sv': 'Edward Elgar', 'fr': 'Edward Elgar', 'da': 'Edward Elgar'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Andy Sachs in 'The Devil Wears Prada'?", 'de': "Welche Schauspielerin spielte Andy Sachs in 'Der Teufel traegt Prada'?", 'es': "¿Que actriz interpreto a Andy Sachs en 'El diablo viste a la moda'?", 'nl': "Welke actrice speelde Andy Sachs in 'The Devil Wears Prada'?", 'sv': "Vilken skadespelerska spelade Andy Sachs i 'Djavulen bar Prada'?", 'fr': "Quelle actrice a joue Andy Sachs dans 'Le Diable s'habille en Prada' ?", 'da': "Hvilken skuespiller spillede Andy Sachs i 'Djævelen går i Prada'?"},
 'opt1': {'en': 'Anne Hathaway', 'de': 'Anne Hathaway', 'es': 'Anne Hathaway', 'nl': 'Anne Hathaway', 'sv': 'Anne Hathaway', 'fr': 'Anne Hathaway', 'da': 'Anne Hathaway'},
 'opt2': {'en': 'Emily Blunt', 'de': 'Emily Blunt', 'es': 'Emily Blunt', 'nl': 'Emily Blunt', 'sv': 'Emily Blunt', 'fr': 'Emily Blunt', 'da': 'Emily Blunt'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Ethiopian flatbread is spongy and used to scoop up stews?', 'de': 'Welches aethiopische Fladenbrot ist schwammig und wird zum Aufnehmen von Eintoepfen verwendet?', 'es': '¿Que pan plano etiope es esponjoso y se usa para recoger guisos?', 'nl': 'Welk Ethiopisch platbrood is sponsachtig en wordt gebruikt om stoofpotten mee op te scheppen?', 'sv': 'Vilket etiopiskt tunnbrod ar svampigt och anvands for att skopa upp grytor?', 'fr': "Quel pain plat ethiopien est spongieux et utilise pour saucer les ragouts ?", 'da': 'Hvilket etiopisk fladbrød er svampet og bruges til at samle gryderetter op med?'},
 'opt1': {'en': 'Injera', 'de': 'Injera', 'es': 'Injera', 'nl': 'Injera', 'sv': 'Injera', 'fr': 'Injera', 'da': 'Injera'},
 'opt2': {'en': 'Naan', 'de': 'Naan', 'es': 'Naan', 'nl': 'Naan', 'sv': 'Naan', 'fr': 'Naan', 'da': 'Naan'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Everglades?', 'de': 'In welchem US-Bundesstaat liegen die Everglades?', 'es': '¿En que estado de EE.UU. estan los Everglades?', 'nl': 'In welke Amerikaanse staat liggen de Everglades?', 'sv': 'I vilken amerikansk delstat ligger Everglades?', 'fr': "Dans quel État americain se trouvent les Everglades ?", 'da': 'I hvilken amerikansk delstat ligger Everglades?'},
 'opt1': {'en': 'Florida', 'de': 'Florida', 'es': 'Florida', 'nl': 'Florida', 'sv': 'Florida', 'fr': 'la Floride', 'da': 'Florida'},
 'opt2': {'en': 'Louisiana', 'de': 'Louisiana', 'es': 'Luisiana', 'nl': 'Louisiana', 'sv': 'Louisiana', 'fr': 'la Louisiane', 'da': 'Louisiana'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large rodent builds dams and lodges in rivers?', 'de': 'Welches grosse Nagetier baut Daemme und Baue in Fluessen?', 'es': '¿Que roedor grande construye presas y madrigueras en los rios?', 'nl': 'Welk groot knaagdier bouwt dammen en burchten in rivieren?', 'sv': 'Vilken stor gnagare bygger dammar och boningar i floder?', 'fr': "Quel grand rongeur construit des barrages et des huttes dans les rivieres ?", 'da': 'Hvilket stort gnaver bygger dæmninger og hytter i floder?'},
 'opt1': {'en': 'Beaver', 'de': 'Biber', 'es': 'Castor', 'nl': 'Bever', 'sv': 'Bavern', 'fr': 'Castor', 'da': 'Bæver'},
 'opt2': {'en': 'Muskrat', 'de': 'Bisamratte', 'es': 'Rata almizclera', 'nl': 'Muskusrat', 'sv': 'Bisamratta', 'fr': 'Rat musque', 'da': 'Bisamrotte'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Japanese festival celebrates cherry blossoms in spring?', 'de': 'Welches japanische Fest feiert die Kirschblueten im Fruehling?', 'es': '¿Que festival japones celebra los cerezos en flor en primavera?', 'nl': 'Welk Japans festival viert de kersenbloesems in de lente?', 'sv': 'Vilken japansk festival firar korsbarsblom pa varen?', 'fr': "Quel festival japonais celebre les fleurs de cerisier au printemps ?", 'da': 'Hvilken japansk festival fejrer kirsebærblomster om foråret?'},
 'opt1': {'en': 'Hanami', 'de': 'Hanami', 'es': 'Hanami', 'nl': 'Hanami', 'sv': 'Hanami', 'fr': 'Hanami', 'da': 'Hanami'},
 'opt2': {'en': 'Obon', 'de': 'Obon', 'es': 'Obon', 'nl': 'Obon', 'sv': 'Obon', 'fr': 'Obon', 'da': 'Obon'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which author wrote 'The Handmaid's Tale'?", 'de': "Welche Autorin schrieb 'Der Report der Magd'?", 'es': "¿Que autora escribio 'El cuento de la criada'?", 'nl': "Welke auteur schreef 'The Handmaid's Tale'?", 'sv': "Vilken forfattare skrev 'Tjanarinnans berattelse'?", 'fr': "Quelle autrice a ecrit 'La Servante ecarlate' ?", 'da': "Hvilken forfatter skrev 'Tjenerindens fortælling'?"},
 'opt1': {'en': 'Margaret Atwood', 'de': 'Margaret Atwood', 'es': 'Margaret Atwood', 'nl': 'Margaret Atwood', 'sv': 'Margaret Atwood', 'fr': 'Margaret Atwood', 'da': 'Margaret Atwood'},
 'opt2': {'en': 'Toni Morrison', 'de': 'Toni Morrison', 'es': 'Toni Morrison', 'nl': 'Toni Morrison', 'sv': 'Toni Morrison', 'fr': 'Toni Morrison', 'da': 'Toni Morrison'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the Blue Lagoon geothermal spa?', 'de': 'In welchem Land liegt das geothermische Spa Blaue Lagune?', 'es': '¿En que pais esta el spa geotermico Laguna Azul?', 'nl': 'In welk land ligt de geothermische spa Blue Lagoon?', 'sv': 'I vilket land ligger den geotermiska span Blue Lagoon?', 'fr': "Dans quel pays se trouve le spa geothermique du Blue Lagoon ?", 'da': 'Hvilket land huser den geotermiske spa Blue Lagoon?'},
 'opt1': {'en': 'Iceland', 'de': 'Island', 'es': 'Islandia', 'nl': 'IJsland', 'sv': 'Island', 'fr': "l'Islande", 'da': 'Island'},
 'opt2': {'en': 'Norway', 'de': 'Norwegen', 'es': 'Noruega', 'nl': 'Noorwegen', 'sv': 'Norge', 'fr': 'la Norvège', 'da': 'Norge'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which instrument is central to flamenco music alongside castanets?', 'de': 'Welches Instrument ist zentral fuer Flamenco-Musik neben Kastagnetten?', 'es': '¿Que instrumento es central en la musica flamenca junto con las castañuelas?', 'nl': 'Welk instrument is centraal in flamencomuziek naast castagnetten?', 'sv': 'Vilket instrument ar centralt i flamencomusik vid sidan av kastanjetter?', 'fr': "Quel instrument est central dans la musique flamenco aux cotes des castagnettes ?", 'da': 'Hvilket instrument er centralt i flamencomusik ved siden af kastagnetter?'},
 'opt1': {'en': 'Spanish guitar', 'de': 'Spanische Gitarre', 'es': 'Guitarra española', 'nl': 'Spaanse gitaar', 'sv': 'Spansk gitarr', 'fr': 'Guitare espagnole', 'da': 'Spansk guitar'},
 'opt2': {'en': 'Mandolin', 'de': 'Mandoline', 'es': 'Mandolina', 'nl': 'Mandoline', 'sv': 'Mandolin', 'fr': 'Mandoline', 'da': 'Mandolin'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a savory casserole made from layers of cheese, noodles, and tomato meat sauce?', 'de': 'Welches suddstaatliche Gericht ist ein herzhafter Auflauf aus Schichten von Kaese, Nudeln und Tomaten-Fleisch-Sauce?', 'es': '¿Que plato sureño de EE.UU. es una cacerola salada hecha de capas de queso, fideos y salsa de tomate con carne?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een hartige ovenschotel gemaakt van lagen kaas, pasta en tomaten-vleessaus?', 'sv': 'Vilken sydstatlig ratt ar en salt gryta gjord av lager ost, pasta och tomatkottsas?', 'fr': "Quel plat du Sud des États-Unis est un gratin sale a base de couches de fromage, de pates et de sauce tomate a la viande ?", 'da': 'Hvilken sydstatsret er en salt gryderet lavet af lag af ost, nudler og tomatkødsovs?'},
 'opt1': {'en': 'Baked ziti', 'de': 'Baked Ziti', 'es': 'Ziti al horno', 'nl': 'Baked ziti', 'sv': 'Ugnsbakad ziti', 'fr': 'Ziti au four', 'da': 'Bagt ziti'},
 'opt2': {'en': 'Lasagna', 'de': 'Lasagne', 'es': 'Lasaña', 'nl': 'Lasagne', 'sv': 'Lasagne', 'fr': 'Lasagne', 'da': 'Lasagne'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which mineral is essential for blood clotting and found in leafy greens?', 'de': 'Welches Vitamin ist essentiell fuer die Blutgerinnung und in Blattgemuese enthalten?', 'es': '¿Que vitamina es esencial para la coagulacion de la sangre y se encuentra en verduras de hoja verde?', 'nl': 'Welk vitamine is essentieel voor bloedstolling en te vinden in bladgroenten?', 'sv': 'Vilket vitamin ar viktigt for blodkoagulering och finns i bladgronsaker?', 'fr': "Quelle vitamine est essentielle a la coagulation sanguine et se trouve dans les legumes-feuilles ?", 'da': 'Hvilket vitamin er essentielt for blodkoagulation og findes i bladgrønt?'},
 'opt1': {'en': 'Vitamin K', 'de': 'Vitamin K', 'es': 'Vitamina K', 'nl': 'Vitamine K', 'sv': 'Vitamin K', 'fr': 'Vitamine K', 'da': 'Vitamin K'},
 'opt2': {'en': 'Vitamin E', 'de': 'Vitamin E', 'es': 'Vitamina E', 'nl': 'Vitamine E', 'sv': 'Vitamin E', 'fr': 'Vitamine E', 'da': 'Vitamin E'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which film has Reese Witherspoon playing a sorority girl who attends Harvard Law?', 'de': 'In welchem Film spielt Reese Witherspoon ein Studentinnenverbindungsmaedchen, das Harvard Law besucht?', 'es': '¿En que pelicula Reese Witherspoon interpreta a una chica de hermandad que asiste a la facultad de derecho de Harvard?', 'nl': 'In welke film speelt Reese Witherspoon een studentenverenigingsmeisje dat rechten studeert aan Harvard?', 'sv': 'I vilken film spelar Reese Witherspoon en studentflicka som gar Harvard Law?', 'fr': "Dans quel film Reese Witherspoon joue-t-elle une etudiante de sororite qui frequente la faculte de droit de Harvard ?", 'da': 'I hvilken film spiller Reese Witherspoon en søsterskabspige, der går på Harvard Law?'},
 'opt1': {'en': 'Legally Blonde', 'de': 'Natuerlich blond', 'es': 'Legalmente rubia', 'nl': 'Legally Blonde', 'sv': 'Blond makt', 'fr': 'La Revanche d\'une blonde', 'da': 'Naturligt blond'},
 'opt2': {'en': 'Sweet Home Alabama', 'de': 'Sweet Home Alabama', 'es': 'Dulce hogar Alabama', 'nl': 'Sweet Home Alabama', 'sv': 'Sweet Home Alabama', 'fr': 'Sweet Home Alabama', 'da': 'Sweet Home Alabama'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which country is home to the ancient city of Petra's neighbor, the Dead Sea?", 'de': 'In welchem Land liegt das Tote Meer, der Nachbar der antiken Stadt Petra?', 'es': '¿En que pais esta el Mar Muerto, vecino de la antigua ciudad de Petra?', 'nl': 'In welk land ligt de Dode Zee, de buur van de oude stad Petra?', 'sv': 'I vilket land ligger Dodahavet, granne till den antika staden Petra?', 'fr': "Dans quel pays se trouve la mer Morte, voisine de l'ancienne cite de Petra ?", 'da': 'Hvilket land huser Det Døde Hav, nabo til den antikke by Petra?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Israel', 'de': 'Israel', 'es': 'Israel', 'nl': 'Israel', 'sv': 'Israel', 'fr': 'Israel', 'da': 'Israel'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of dress, popularized in the 1950s, features a fitted bodice and full skirt?', 'de': 'Welche Art von Kleid, populaer in den 1950er Jahren, hat ein taillierten Oberteil und einen weiten Rock?', 'es': '¿Que tipo de vestido, popularizado en los años 50, presenta un corpiño ajustado y falda amplia?', 'nl': 'Welk type jurk, populair in de jaren 50, heeft een getailleerd lijfje en een wijde rok?', 'sv': 'Vilken typ av klanning, popular pa 1950-talet, har en atsittande liv och full kjol?', 'fr': "Quel type de robe, popularise dans les annees 1950, presente un corsage ajuste et une jupe ample ?", 'da': 'Hvilken type kjole, populariseret i 1950’erne, har en tætsiddende overdel og fuld nederdel?'},
 'opt1': {'en': 'A-line dress', 'de': 'A-Linien-Kleid', 'es': 'Vestido en linea A', 'nl': 'A-lijn jurk', 'sv': 'A-linjeklanning', 'fr': 'Robe trapeze', 'da': 'A-linje kjole'},
 'opt2': {'en': 'Sheath dress', 'de': 'Etuikleid', 'es': 'Vestido tubo', 'nl': 'Kokerjurk', 'sv': 'Etuiklanning', 'fr': 'Robe fourreau', 'da': 'Etuikjole'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which venomous snake is famous for the warning rattle at the end of its tail?', 'de': 'Welche Giftschlange ist beruehmt fuer die Warnklapper am Ende ihres Schwanzes?', 'es': '¿Que serpiente venenosa es famosa por el cascabel de advertencia al final de su cola?', 'nl': 'Welke giftige slang staat bekend om de waarschuwingsratel aan het einde van zijn staart?', 'sv': 'Vilken giftorm ar kand for varningsskallran i slutet av sin svans?', 'fr': "Quel serpent venimeux est celebre pour son hochet d'avertissement au bout de la queue ?", 'da': 'Hvilken giftig slange er berømt for advarselsraslen for enden af sin hale?'},
 'opt1': {'en': 'Rattlesnake', 'de': 'Klapperschlange', 'es': 'Serpiente de cascabel', 'nl': 'Ratelslang', 'sv': 'Skallerorm', 'fr': 'Serpent a sonnette', 'da': 'Klapperslange'},
 'opt2': {'en': 'Cobra', 'de': 'Kobra', 'es': 'Cobra', 'nl': 'Cobra', 'sv': 'Kobra', 'fr': 'Cobra', 'da': 'Kobra'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady hosted the first televised White House tour?', 'de': 'Welche First Lady veranstaltete die erste im Fernsehen uebertragene Fuehrung durchs Weisse Haus?', 'es': '¿Que primera dama presento la primera visita televisada de la Casa Blanca?', 'nl': 'Welke First Lady presenteerde de eerste televisie-rondleiding door het Witte Huis?', 'sv': 'Vilken First Lady var vard for den forsta TV-sanda rundturen i Vita huset?', 'fr': "Quelle Première dame a anime la premiere visite televisee de la Maison-Blanche ?", 'da': 'Hvilken førstedame var vært for den første tv-transmitterede rundvisning i Det Hvide Hus?'},
 'opt1': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'},
 'opt2': {'en': 'Mamie Eisenhower', 'de': 'Mamie Eisenhower', 'es': 'Mamie Eisenhower', 'nl': 'Mamie Eisenhower', 'sv': 'Mamie Eisenhower', 'fr': 'Mamie Eisenhower', 'da': 'Mamie Eisenhower'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which First Lady was a published author of children\'s books about the White House?', 'de': 'Welche First Lady war eine veroeffentlichte Autorin von Kinderbuechern ueber das Weisse Haus?', 'es': '¿Que primera dama fue autora publicada de libros infantiles sobre la Casa Blanca?', 'nl': 'Welke First Lady was een gepubliceerde auteur van kinderboeken over het Witte Huis?', 'sv': 'Vilken First Lady var en publicerad forfattare av barnbocker om Vita huset?', 'fr': "Quelle Première dame etait autrice publiee de livres pour enfants sur la Maison-Blanche ?", 'da': 'Hvilken førstedame var en udgivet forfatter af børnebøger om Det Hvide Hus?'},
 'opt1': {'en': 'Laura Bush', 'de': 'Laura Bush', 'es': 'Laura Bush', 'nl': 'Laura Bush', 'sv': 'Laura Bush', 'fr': 'Laura Bush', 'da': 'Laura Bush'},
 'opt2': {'en': 'Barbara Bush', 'de': 'Barbara Bush', 'es': 'Barbara Bush', 'nl': 'Barbara Bush', 'sv': 'Barbara Bush', 'fr': 'Barbara Bush', 'da': 'Barbara Bush'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Colombian author wrote 'One Hundred Years of Solitude'?", 'de': "Welcher kolumbianische Autor schrieb 'Hundert Jahre Einsamkeit'?", 'es': "¿Que autor colombiano escribio 'Cien años de soledad'?", 'nl': "Welke Colombiaanse auteur schreef 'Honderd jaar eenzaamheid'?", 'sv': "Vilken colombiansk forfattare skrev 'Hundra ars ensamhet'?", 'fr': "Quel auteur colombien a ecrit 'Cent ans de solitude' ?", 'da': "Hvilken colombiansk forfatter skrev 'Hundrede års ensomhed'?"},
 'opt1': {'en': 'Gabriel Garcia Marquez', 'de': 'Gabriel Garcia Marquez', 'es': 'Gabriel Garcia Marquez', 'nl': 'Gabriel Garcia Marquez', 'sv': 'Gabriel Garcia Marquez', 'fr': 'Gabriel Garcia Marquez', 'da': 'Gabriel Garcia Marquez'},
 'opt2': {'en': 'Isabel Allende', 'de': 'Isabel Allende', 'es': 'Isabel Allende', 'nl': 'Isabel Allende', 'sv': 'Isabel Allende', 'fr': 'Isabel Allende', 'da': 'Isabel Allende'}},

{'category': 'Science', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which vitamin is produced by gut bacteria and aids blood clotting?', 'de': 'Welches Vitamin wird von Darmbakterien produziert und unterstuetzt die Blutgerinnung?', 'es': '¿Que vitamina es producida por bacterias intestinales y ayuda a la coagulacion de la sangre?', 'nl': 'Welk vitamine wordt geproduceerd door darmbacterien en helpt bij bloedstolling?', 'sv': 'Vilket vitamin produceras av tarmbakterier och hjalper blodkoagulering?', 'fr': "Quelle vitamine est produite par les bacteries intestinales et aide la coagulation sanguine ?", 'da': 'Hvilket vitamin produceres af tarmbakterier og hjælper blodkoagulation?'},
 'opt1': {'en': 'Vitamin K', 'de': 'Vitamin K', 'es': 'Vitamina K', 'nl': 'Vitamine K', 'sv': 'Vitamin K', 'fr': 'Vitamine K', 'da': 'Vitamin K'},
 'opt2': {'en': 'Vitamin B12', 'de': 'Vitamin B12', 'es': 'Vitamina B12', 'nl': 'Vitamine B12', 'sv': 'Vitamin B12', 'fr': 'Vitamine B12', 'da': 'Vitamin B12'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Spanish painter created 'Guernica' in response to the Spanish Civil War?", 'de': "Welcher spanische Maler schuf 'Guernica' als Reaktion auf den Spanischen Buergerkrieg?", 'es': "¿Que pintor español creo 'Guernica' en respuesta a la Guerra Civil Española?", 'nl': "Welke Spaanse schilder creëerde 'Guernica' als reactie op de Spaanse Burgeroorlog?", 'sv': "Vilken spansk malare skapade 'Guernica' som svar pa spanska inbordeskriget?", 'fr': "Quel peintre espagnol a cree 'Guernica' en reponse a la guerre civile espagnole ?", 'da': "Hvilken spansk maler skabte 'Guernica' som svar på den spanske borgerkrig?"},
 'opt1': {'en': 'Pablo Picasso', 'de': 'Pablo Picasso', 'es': 'Pablo Picasso', 'nl': 'Pablo Picasso', 'sv': 'Pablo Picasso', 'fr': 'Pablo Picasso', 'da': 'Pablo Picasso'},
 'opt2': {'en': 'Joan Miro', 'de': 'Joan Miro', 'es': 'Joan Miro', 'nl': 'Joan Miro', 'sv': 'Joan Miro', 'fr': 'Joan Miro', 'da': 'Joan Miro'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Southeast Asian country is made up of over 17,000 islands?', 'de': 'Welches suedostasiatische Land besteht aus ueber 17.000 Inseln?', 'es': '¿Que pais del sudeste asiatico esta formado por mas de 17.000 islas?', 'nl': 'Welk Zuidoost-Aziatisch land bestaat uit meer dan 17.000 eilanden?', 'sv': 'Vilket sydostasiatiskt land bestar av over 17 000 oar?', 'fr': "Quel pays d'Asie du Sud-Est est compose de plus de 17 000 iles ?", 'da': 'Hvilket sydøstasiatisk land består af over 17.000 øer?'},
 'opt1': {'en': 'Indonesia', 'de': 'Indonesien', 'es': 'Indonesia', 'nl': 'Indonesie', 'sv': 'Indonesien', 'fr': "l'Indonesie", 'da': 'Indonesien'},
 'opt2': {'en': 'Philippines', 'de': 'Philippinen', 'es': 'Filipinas', 'nl': 'Filipijnen', 'sv': 'Filippinerna', 'fr': 'les Philippines', 'da': 'Filippinerne'}},

{'category': 'Music', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Austrian composer wrote 'The Marriage of Figaro'?", 'de': "Welcher oesterreichische Komponist schrieb 'Die Hochzeit des Figaro'?", 'es': "¿Que compositor austriaco escribio 'Las bodas de Figaro'?", 'nl': "Welke Oostenrijkse componist schreef 'De bruiloft van Figaro'?", 'sv': "Vilken osterrikisk kompositor skrev 'Figaros brollop'?", 'fr': "Quel compositeur autrichien a ecrit 'Les Noces de Figaro' ?", 'da': "Hvilken østrigsk komponist skrev 'Figaros bryllup'?"},
 'opt1': {'en': 'Wolfgang Amadeus Mozart', 'de': 'Wolfgang Amadeus Mozart', 'es': 'Wolfgang Amadeus Mozart', 'nl': 'Wolfgang Amadeus Mozart', 'sv': 'Wolfgang Amadeus Mozart', 'fr': 'Wolfgang Amadeus Mozart', 'da': 'Wolfgang Amadeus Mozart'},
 'opt2': {'en': 'Franz Schubert', 'de': 'Franz Schubert', 'es': 'Franz Schubert', 'nl': 'Franz Schubert', 'sv': 'Franz Schubert', 'fr': 'Franz Schubert', 'da': 'Franz Schubert'}},

{'category': 'Food', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which French dish consists of simmered beef in red wine with mushrooms and onions?', 'de': 'Welches franzoesische Gericht besteht aus geschmortem Rindfleisch in Rotwein mit Pilzen und Zwiebeln?', 'es': '¿Que plato frances consiste en carne de res cocida a fuego lento en vino tinto con champiñones y cebollas?', 'nl': 'Welk Frans gerecht bestaat uit gestoofd rundvlees in rode wijn met champignons en uien?', 'sv': 'Vilken fransk ratt bestar av kott smaksatt i rodvin med svamp och lok?', 'fr': "Quel plat français consiste en du boeuf mijote au vin rouge avec des champignons et des oignons ?", 'da': 'Hvilken fransk ret består af oksekød simret i rødvin med svampe og løg?'},
 'opt1': {'en': 'Boeuf bourguignon', 'de': 'Boeuf Bourguignon', 'es': 'Boeuf bourguignon', 'nl': 'Boeuf bourguignon', 'sv': 'Boeuf bourguignon', 'fr': 'Boeuf bourguignon', 'da': 'Boeuf bourguignon'},
 'opt2': {'en': 'Coq au vin', 'de': 'Coq au vin', 'es': 'Coq au vin', 'nl': 'Coq au vin', 'sv': 'Coq au vin', 'fr': 'Coq au vin', 'da': 'Coq au vin'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Norse mythology, which one-eyed god is the ruler of Asgard?', 'de': 'In der nordischen Mythologie, welcher einaeugige Gott ist Herrscher von Asgard?', 'es': '¿En la mitologia nordica, que dios tuerto es el gobernante de Asgard?', 'nl': 'In de Noorse mythologie, welke eenogige god is de heerser van Asgard?', 'sv': 'I nordisk mytologi, vilken enogd gud ar harskaren over Asgard?', 'fr': "Dans la mythologie nordique, quel dieu borgne est le souverain d'Asgard ?", 'da': 'I nordisk mytologi, hvilken enøjet gud er hersker over Asgård?'},
 'opt1': {'en': 'Odin', 'de': 'Odin', 'es': 'Odin', 'nl': 'Odin', 'sv': 'Oden', 'fr': 'Odin', 'da': 'Odin'},
 'opt2': {'en': 'Thor', 'de': 'Thor', 'es': 'Thor', 'nl': 'Thor', 'sv': 'Tor', 'fr': 'Thor', 'da': 'Thor'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which actress won an Oscar for playing a deaf woman in 'Children of a Lesser God'?", 'de': "Welche Schauspielerin gewann einen Oscar fuer ihre Rolle als gehoerlose Frau in 'Gottes vergessene Kinder'?", 'es': "¿Que actriz gano un Oscar por interpretar a una mujer sorda en 'Hijos de un Dios menor'?", 'nl': "Welke actrice won een Oscar voor haar rol als dove vrouw in 'Children of a Lesser God'?", 'sv': "Vilken skadespelerska vann en Oscar for att ha spelat en dov kvinna i 'Guds glomda barn'?", 'fr': "Quelle actrice a remporte un Oscar pour son role de femme sourde dans 'Les Enfants du silence' ?", 'da': "Hvilken skuespiller vandt en Oscar for at spille en døv kvinde i 'Guds glemte børn'?"},
 'opt1': {'en': 'Marlee Matlin', 'de': 'Marlee Matlin', 'es': 'Marlee Matlin', 'nl': 'Marlee Matlin', 'sv': 'Marlee Matlin', 'fr': 'Marlee Matlin', 'da': 'Marlee Matlin'},
 'opt2': {'en': 'Holly Hunter', 'de': 'Holly Hunter', 'es': 'Holly Hunter', 'nl': 'Holly Hunter', 'sv': 'Holly Hunter', 'fr': 'Holly Hunter', 'da': 'Holly Hunter'}},

{'category': 'Space', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which NASA mission famously had to abort its moon landing after an oxygen tank explosion, safely returning the crew to Earth?', 'de': 'Welche NASA-Mission musste ihre Mondlandung nach einer Sauerstofftank-Explosion beruehmt abbrechen und die Besatzung sicher zur Erde zurueckbringen?', 'es': '¿Que mision de la NASA tuvo que abortar famosamente su alunizaje tras una explosion de un tanque de oxigeno, devolviendo a salvo a la tripulacion a la Tierra?', 'nl': 'Welke NASA-missie moest beroemd de maanlanding afbreken na een zuurstoftank-explosie en bracht de bemanning veilig terug naar de aarde?', 'sv': 'Vilket NASA-uppdrag fick beromt avbryta sin manlandning efter en syretanksexplosion, och fick besattningen sakert tillbaka till jorden?', 'fr': "Quelle mission de la NASA a du celebrement annuler son alunissage apres l'explosion d'un reservoir d'oxygene, ramenant l'equipage sain et sauf sur Terre ?", 'da': 'Hvilken NASA-mission måtte berømt afbryde sin månelanding efter en iltbeholder-eksplosion og bragte besætningen sikkert tilbage til Jorden?'},
 'opt1': {'en': 'Apollo 13', 'de': 'Apollo 13', 'es': 'Apolo 13', 'nl': 'Apollo 13', 'sv': 'Apollo 13', 'fr': 'Apollo 13', 'da': 'Apollo 13'},
 'opt2': {'en': 'Apollo 11', 'de': 'Apollo 11', 'es': 'Apolo 11', 'nl': 'Apollo 11', 'sv': 'Apollo 11', 'fr': 'Apollo 11', 'da': 'Apollo 11'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient Nazca Lines?', 'de': 'In welchem Land liegen die antiken Nazca-Linien?', 'es': '¿En que pais estan las antiguas Lineas de Nazca?', 'nl': 'In welk land liggen de oude Nazca-lijnen?', 'sv': 'I vilket land ligger de antika Nazca-linjerna?', 'fr': "Dans quel pays se trouvent les anciennes lignes de Nazca ?", 'da': 'Hvilket land huser de gamle Nazca-linjer?'},
 'opt1': {'en': 'Peru', 'de': 'Peru', 'es': 'Peru', 'nl': 'Peru', 'sv': 'Peru', 'fr': 'le Perou', 'da': 'Peru'},
 'opt2': {'en': 'Chile', 'de': 'Chile', 'es': 'Chile', 'nl': 'Chili', 'sv': 'Chile', 'fr': 'le Chili', 'da': 'Chile'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which fashion designer is credited with introducing the 'New Look' silhouette after WWII?", 'de': "Welcher Modedesigner wird die Einfuehrung der 'New Look'-Silhouette nach dem Zweiten Weltkrieg zugeschrieben?", 'es': "¿A que diseñador de moda se le atribuye la introduccion de la silueta 'New Look' despues de la Segunda Guerra Mundial?", 'nl': "Welke modeontwerper wordt gecrediteerd met de introductie van de 'New Look'-silhouet na WOII?", 'sv': "Vilken modeskapare tillskrivs att ha introducerat 'New Look'-silhuetten efter andra varldskriget?", 'fr': "Quel createur de mode est credite d'avoir introduit la silhouette 'New Look' apres la Seconde Guerre mondiale ?", 'da': "Hvilken modedesigner krediteres for at have introduceret 'New Look'-silhuetten efter 2. verdenskrig?"},
 'opt1': {'en': 'Christian Dior', 'de': 'Christian Dior', 'es': 'Christian Dior', 'nl': 'Christian Dior', 'sv': 'Christian Dior', 'fr': 'Christian Dior', 'da': 'Christian Dior'},
 'opt2': {'en': 'Coco Chanel', 'de': 'Coco Chanel', 'es': 'Coco Chanel', 'nl': 'Coco Chanel', 'sv': 'Coco Chanel', 'fr': 'Coco Chanel', 'da': 'Coco Chanel'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Chinese philosopher's teachings emphasize filial piety and social harmony?", 'de': 'Welcher chinesische Philosoph betont in seinen Lehren kindliche Pietaet und soziale Harmonie?', 'es': '¿Que filosofo chino enfatiza en sus enseñanzas la piedad filial y la armonia social?', 'nl': 'Welke Chinese filosoof benadrukt in zijn leer kinderlijke vroomheid en sociale harmonie?', 'sv': 'Vilken kinesisk filosofs laror betonar barnslig fromhet och social harmoni?', 'fr': "Quel philosophe chinois met l'accent sur la piete filiale et l'harmonie sociale dans ses enseignements ?", 'da': 'Hvilken kinesisk filosofs lære lægger vægt på filial fromhed og social harmoni?'},
 'opt1': {'en': 'Confucius', 'de': 'Konfuzius', 'es': 'Confucio', 'nl': 'Confucius', 'sv': 'Konfucius', 'fr': 'Confucius', 'da': 'Konfucius'},
 'opt2': {'en': 'Laozi', 'de': 'Laozi', 'es': 'Laozi', 'nl': 'Laozi', 'sv': 'Laozi', 'fr': 'Laozi', 'da': 'Laozi'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which quilting pattern uses hexagon shapes pieced together, popular in English paper piecing?', 'de': 'Welches Quiltmuster verwendet zusammengesetzte Sechseckformen, beliebt beim English Paper Piecing?', 'es': '¿Que patron de acolchado usa formas hexagonales unidas, popular en el paper piecing ingles?', 'nl': 'Welk quiltpatroon gebruikt in elkaar gezette zeshoekvormen, populair bij Engelse papiervoering?', 'sv': 'Vilket lapptacksmonster anvander hopsatta sexkantsformer, populart i engelsk pappersflismetod?', 'fr': "Quel motif de patchwork utilise des formes hexagonales assemblees, populaire dans le paper piecing anglais ?", 'da': 'Hvilket quiltmønster bruger sammensatte sekskantformer, populært i engelsk papirpiecing?'},
 'opt1': {'en': "Grandmother's Flower Garden", 'de': 'Großmutters Blumengarten', 'es': 'Jardin de flores de la abuela', 'nl': 'Grootmoeders Bloementuin', 'sv': 'Farmors blomstertradgard', 'fr': 'Le jardin de fleurs de grand-mere', 'da': 'Bedstemors Blomsterhave'},
 'opt2': {'en': 'Log Cabin', 'de': 'Blockhuette', 'es': 'Cabaña de troncos', 'nl': 'Blokhut', 'sv': 'Timmerstuga', 'fr': 'Cabane en rondins', 'da': 'Log Cabin'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which flowering vine is known for its fragrant white blooms that open at night?', 'de': 'Welche bluehende Kletterpflanze ist fuer ihre duftenden weissen Blueten bekannt, die sich nachts oeffnen?', 'es': '¿Que enredadera es conocida por sus fragantes flores blancas que se abren de noche?', 'nl': 'Welke bloeiende klimplant staat bekend om zijn geurige witte bloemen die \'s nachts opengaan?', 'sv': 'Vilken blommande klattervaxt ar kand for sina doftande vita blommor som oppnar sig pa natten?', 'fr': "Quelle liane a fleurs est connue pour ses fleurs blanches parfumees qui s'ouvrent la nuit ?", 'da': 'Hvilken blomstrende slyngplante er kendt for sine duftende hvide blomster, der åbner om natten?'},
 'opt1': {'en': 'Moonflower', 'de': 'Mondblume', 'es': 'Flor de luna', 'nl': 'Maanbloem', 'sv': 'Manblomma', 'fr': 'Fleur de lune', 'da': 'Måneblomst'},
 'opt2': {'en': 'Morning glory', 'de': 'Prunkwinde', 'es': 'Campanilla', 'nl': 'Winde', 'sv': 'Marvinda', 'fr': 'Belle-de-jour', 'da': 'Snerle'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which architect designed the Sydney Opera House with its distinctive sail-like roofs?', 'de': 'Welcher Architekt entwarf das Sydney Opera House mit seinen markanten segelaehnlichen Daechern?', 'es': '¿Que arquitecto diseño la Opera de Sidney con sus distintivos techos en forma de vela?', 'nl': 'Welke architect ontwierp het Sydney Opera House met zijn kenmerkende zeilvormige daken?', 'sv': 'Vilken arkitekt designade Sydneys operahus med sina utmarkande segelliknande tak?', 'fr': "Quel architecte a conçu l'Opera de Sydney avec ses toits caracteristiques en forme de voile ?", 'da': 'Hvilken arkitekt designede Sydney Opera House med dets karakteristiske sejlformede tage?'},
 'opt1': {'en': 'Jorn Utzon', 'de': 'Jorn Utzon', 'es': 'Jorn Utzon', 'nl': 'Jorn Utzon', 'sv': 'Jorn Utzon', 'fr': 'Jorn Utzon', 'da': 'Jørn Utzon'},
 'opt2': {'en': 'Renzo Piano', 'de': 'Renzo Piano', 'es': 'Renzo Piano', 'nl': 'Renzo Piano', 'sv': 'Renzo Piano', 'fr': 'Renzo Piano', 'da': 'Renzo Piano'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Norse mythology, which trickster god is known for shapeshifting?', 'de': 'In der nordischen Mythologie, welcher Schelmgott ist fuer Gestaltwandel bekannt?', 'es': '¿En la mitologia nordica, que dios embaucador es conocido por cambiar de forma?', 'nl': 'In de Noorse mythologie, welke bedriegersgod staat bekend om gedaanteverwisseling?', 'sv': 'I nordisk mytologi, vilken lurgud ar kand for skepnadsskifte?', 'fr': "Dans la mythologie nordique, quel dieu farceur est connu pour se metamorphoser ?", 'da': 'I nordisk mytologi, hvilken narreri-gud er kendt for at skifte skikkelse?'},
 'opt1': {'en': 'Loki', 'de': 'Loki', 'es': 'Loki', 'nl': 'Loki', 'sv': 'Loke', 'fr': 'Loki', 'da': 'Loke'},
 'opt2': {'en': 'Baldr', 'de': 'Baldur', 'es': 'Baldr', 'nl': 'Baldr', 'sv': 'Balder', 'fr': 'Baldr', 'da': 'Balder'}},
]
