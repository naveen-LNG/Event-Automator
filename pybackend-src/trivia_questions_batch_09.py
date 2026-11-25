# -*- coding: utf-8 -*-
# Trivia Challenge batch #9. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which brown legume is used to make traditional baked beans?', 'de': 'Welche braune Huelsenfrucht wird fuer traditionelle gebackene Bohnen verwendet?', 'es': '¿Que legumbre marron se usa para hacer los frijoles horneados tradicionales?', 'nl': 'Welke bruine peulvrucht wordt gebruikt om traditionele gebakken bonen te maken?', 'sv': 'Vilken brun balja anvands for att gora traditionella bakade bonor?', 'fr': "Quel legume brun est utilise pour faire les haricots en sauce traditionnels ?", 'da': 'Hvilken brun bælgfrugt bruges til at lave traditionelle bagte bønner?'},
 'opt1': {'en': 'Kidney bean', 'de': 'Kidneybohne', 'es': 'Frijol rojo', 'nl': 'Kidneyboon', 'sv': 'Rda bonan', 'fr': 'Le haricot rouge', 'da': 'Kidneybønne'},
 'opt2': {'en': 'Navy bean', 'de': 'Weisse Bohne', 'es': 'Alubia blanca', 'nl': 'Witte boon', 'sv': 'Vit bona', 'fr': 'Le haricot blanc', 'da': 'Hvid bønne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Damascus, one of the oldest continuously inhabited cities?', 'de': 'In welchem Land liegt die antike Stadt Damaskus, eine der aeltesten ununterbrochen bewohnten Staedte?', 'es': '¿En que pais esta la antigua ciudad de Damasco, una de las ciudades habitadas continuamente mas antiguas?', 'nl': 'In welk land ligt de oude stad Damascus, een van de oudste ononderbroken bewoonde steden?', 'sv': 'I vilket land ligger den antika staden Damaskus, en av de aldsta kontinuerligt bebodda staderna?', 'fr': "Dans quel pays se trouve l'ancienne cite de Damas, l'une des plus anciennes villes continuellement habitees ?", 'da': 'Hvilket land huser den antikke by Damaskus, en af de ældste uafbrudt beboede byer?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Syria', 'de': 'Syrien', 'es': 'Siria', 'nl': 'Syrie', 'sv': 'Syrien', 'fr': 'la Syrie', 'da': 'Syrien'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which classic film features a boy who doesn't want to grow up and flies to Neverland?", 'de': 'In welchem Filmklassiker fliegt ein Junge, der nicht erwachsen werden will, ins Nimmerland?', 'es': '¿Que pelicula clasica presenta a un niño que no quiere crecer y vuela a Nunca Jamas?', 'nl': 'In welke klassieke film vliegt een jongen die niet wil opgroeien naar Nooitgedachtland?', 'sv': 'I vilken klassisk film flyger en pojke som inte vill vaxa upp till Landet Ingenstans?', 'fr': "Quel film classique met en scene un garçon qui ne veut pas grandir et s'envole vers le Pays imaginaire ?", 'da': 'Hvilken klassisk film har en dreng, der ikke vil blive voksen, og flyver til Neverland?'},
 'opt1': {'en': 'Hook', 'de': 'Hook', 'es': 'Hook', 'nl': 'Hook', 'sv': 'Hook', 'fr': 'Hook', 'da': 'Hook'},
 'opt2': {'en': 'Peter Pan', 'de': 'Peter Pan', 'es': 'Peter Pan', 'nl': 'Peter Pan', 'sv': 'Peter Pan', 'fr': 'Peter Pan', 'da': 'Peter Pan'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which pet is known for its loyalty and is often trained as a service animal?', 'de': 'Welches Haustier ist fuer seine Treue bekannt und wird oft als Servicetier ausgebildet?', 'es': '¿Que mascota es conocida por su lealtad y a menudo se entrena como animal de servicio?', 'nl': 'Welk huisdier staat bekend om zijn loyaliteit en wordt vaak getraind als hulphond?', 'sv': 'Vilket sallskapsdjur ar kant for sin lojalitet och trans ofta som tjanstedjur?', 'fr': "Quel animal de compagnie est connu pour sa loyaute et est souvent dresse comme animal d'assistance ?", 'da': 'Hvilket kæledyr er kendt for sin loyalitet og trænes ofte som servicedyr?'},
 'opt1': {'en': 'Dog', 'de': 'Hund', 'es': 'Perro', 'nl': 'Hond', 'sv': 'Hund', 'fr': 'Chien', 'da': 'Hund'},
 'opt2': {'en': 'Cat', 'de': 'Katze', 'es': 'Gato', 'nl': 'Kat', 'sv': 'Katt', 'fr': 'Chat', 'da': 'Kat'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which yellow fruit is peeled before eating and is a good source of potassium?', 'de': 'Welche gelbe Frucht wird vor dem Essen geschaelt und ist eine gute Kaliumquelle?', 'es': '¿Que fruta amarilla se pela antes de comer y es una buena fuente de potasio?', 'nl': 'Welk geel fruit wordt geschild voor het eten en is een goede bron van kalium?', 'sv': 'Vilken gul frukt skalas fore atande och ar en bra kaliumkalla?', 'fr': "Quel fruit jaune se pele avant d'etre mange et est une bonne source de potassium ?", 'da': 'Hvilken gul frugt skrælles før spisning og er en god kilde til kalium?'},
 'opt1': {'en': 'Mango', 'de': 'Mango', 'es': 'Mango', 'nl': 'Mango', 'sv': 'Mango', 'fr': 'Mangue', 'da': 'Mango'},
 'opt2': {'en': 'Banana', 'de': 'Banane', 'es': 'Platano', 'nl': 'Banaan', 'sv': 'Banan', 'fr': 'Banane', 'da': 'Banan'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the Great Ocean Road, a famous coastal driving route?', 'de': 'In welchem Land liegt die Great Ocean Road, eine beruehmte Kuestenstrasse?', 'es': '¿En que pais esta la Great Ocean Road, una famosa ruta costera?', 'nl': 'In welk land ligt de Great Ocean Road, een beroemde kustweg?', 'sv': 'I vilket land ligger Great Ocean Road, en beromd kustvag?', 'fr': "Dans quel pays se trouve la Great Ocean Road, une celebre route cotiere ?", 'da': 'Hvilket land huser Great Ocean Road, en berømt kystvej?'},
 'opt1': {'en': 'New Zealand', 'de': 'Neuseeland', 'es': 'Nueva Zelanda', 'nl': 'Nieuw-Zeeland', 'sv': 'Nya Zeeland', 'fr': 'la Nouvelle-Zelande', 'da': 'New Zealand'},
 'opt2': {'en': 'Australia', 'de': 'Australien', 'es': 'Australia', 'nl': 'Australie', 'sv': 'Australien', 'fr': "l'Australie", 'da': 'Australien'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. president is featured on the five-dollar bill?', 'de': 'Welcher US-Praesident ist auf der Fuenf-Dollar-Note abgebildet?', 'es': '¿Que presidente de EE.UU. aparece en el billete de cinco dolares?', 'nl': 'Welke Amerikaanse president staat afgebeeld op het vijf-dollarbiljet?', 'sv': 'Vilken amerikansk president avbildas pa femdollarsedeln?', 'fr': "Quel president americain figure sur le billet de cinq dollars ?", 'da': 'Hvilken amerikansk præsident er afbildet på femdollarsedlen?'},
 'opt1': {'en': 'Abraham Lincoln', 'de': 'Abraham Lincoln', 'es': 'Abraham Lincoln', 'nl': 'Abraham Lincoln', 'sv': 'Abraham Lincoln', 'fr': 'Abraham Lincoln', 'da': 'Abraham Lincoln'},
 'opt2': {'en': 'Ulysses S. Grant', 'de': 'Ulysses S. Grant', 'es': 'Ulysses S. Grant', 'nl': 'Ulysses S. Grant', 'sv': 'Ulysses S. Grant', 'fr': 'Ulysses S. Grant', 'da': 'Ulysses S. Grant'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves cutting and pasting paper or fabric onto a surface to create designs?', 'de': 'Bei welchem Handwerk werden Papier oder Stoff geschnitten und auf eine Oberflaeche geklebt, um Muster zu erstellen?', 'es': '¿Que manualidad implica cortar y pegar papel o tela sobre una superficie para crear diseños?', 'nl': 'Bij welk handwerk knip en plak je papier of stof op een oppervlak om patronen te maken?', 'sv': 'Vilket hantverk innebar att klippa och klistra papper eller tyg pa en yta for att skapa monster?', 'fr': "Quel art manuel consiste a decouper et coller du papier ou du tissu sur une surface pour creer des motifs ?", 'da': 'Hvilket håndværk involverer at klippe og lime papir eller stof på en overflade for at skabe mønstre?'},
 'opt1': {'en': 'Collage', 'de': 'Collage', 'es': 'Collage', 'nl': 'Collage', 'sv': 'Collage', 'fr': 'Le collage', 'da': 'Collage'},
 'opt2': {'en': 'Decoupage', 'de': 'Decoupage', 'es': 'Decoupage', 'nl': 'Decoupage', 'sv': 'Decoupage', 'fr': 'Le decoupage', 'da': 'Decoupage'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which percussion instrument consists of a set of drums played by a single musician?', 'de': 'Welches Perkussionsinstrument besteht aus einem Satz Trommeln, die von einem einzigen Musiker gespielt werden?', 'es': '¿Que instrumento de percusion consiste en un conjunto de tambores tocados por un solo musico?', 'nl': 'Welk slaginstrument bestaat uit een set trommels die door een enkele muzikant worden bespeeld?', 'sv': 'Vilket slaginstrument bestar av en uppsattning trummor som spelas av en ensam musiker?', 'fr': "Quel instrument de percussion consiste en un ensemble de tambours joues par un seul musicien ?", 'da': 'Hvilket slagtøjsinstrument består af et sæt trommer spillet af en enkelt musiker?'},
 'opt1': {'en': 'Drum kit', 'de': 'Schlagzeug', 'es': 'Bateria', 'nl': 'Drumstel', 'sv': 'Trumset', 'fr': 'Batterie', 'da': 'Trommesæt'},
 'opt2': {'en': 'Bongo drums', 'de': 'Bongo-Trommeln', 'es': 'Bongos', 'nl': 'Bongotrommels', 'sv': 'Bongotrummor', 'fr': 'Bongos', 'da': 'Bongotrommer'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which natural phenomenon creates a colorful arc in the sky after rain?', 'de': 'Welches Naturphaenomen erzeugt nach dem Regen einen bunten Bogen am Himmel?', 'es': '¿Que fenomeno natural crea un arco colorido en el cielo despues de la lluvia?', 'nl': 'Welk natuurverschijnsel creëert een kleurrijke boog aan de hemel na regen?', 'sv': 'Vilket naturfenomen skapar en fargglad bage pa himlen efter regn?', 'fr': "Quel phenomene naturel cree un arc colore dans le ciel apres la pluie ?", 'da': 'Hvilket naturfænomen skaber en farverig bue på himlen efter regn?'},
 'opt1': {'en': 'Rainbow', 'de': 'Regenbogen', 'es': 'Arcoiris', 'nl': 'Regenboog', 'sv': 'Regnbage', 'fr': 'Arc-en-ciel', 'da': 'Regnbue'},
 'opt2': {'en': 'Aurora', 'de': 'Polarlicht', 'es': 'Aurora', 'nl': 'Poollicht', 'sv': 'Norrsken', 'fr': 'Aurore boreale', 'da': 'Nordlys'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic film features a family of superheroes with powers like invisibility and stretching?', 'de': 'In welchem Filmklassiker kommt eine Familie von Superhelden mit Kraeften wie Unsichtbarkeit und Dehnbarkeit vor?', 'es': '¿Que pelicula clasica presenta a una familia de superheroes con poderes como invisibilidad y elasticidad?', 'nl': 'In welke klassieke film komt een familie van superhelden voor met krachten zoals onzichtbaarheid en rekbaarheid?', 'sv': 'I vilken klassisk film forekommer en superhjaltefamilj med krafter som osynlighet och tojbarhet?', 'fr': "Quel film classique met en scene une famille de super-heros avec des pouvoirs comme l'invisibilite et l'elasticite ?", 'da': 'Hvilken klassisk film har en superheltefamilie med kræfter som usynlighed og strækning?'},
 'opt1': {'en': 'Sky High', 'de': 'Sky High', 'es': 'Sky High', 'nl': 'Sky High', 'sv': 'Sky High', 'fr': 'Sky High', 'da': 'Sky High'},
 'opt2': {'en': 'The Incredibles', 'de': 'Die Unglaublichen', 'es': 'Los Increibles', 'nl': 'The Incredibles', 'sv': 'Superfamiljen', 'fr': 'Les Indestructibles', 'da': 'Superfamilien'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which U.S. state is known as the 'Silver State'?", 'de': "Welcher US-Bundesstaat ist als 'Silver State' bekannt?", 'es': "¿Que estado de EE.UU. es conocido como el 'Estado de la Plata'?", 'nl': "Welke Amerikaanse staat staat bekend als de 'Silver State'?", 'sv': "Vilken amerikansk delstat ar kand som 'Silverstaten'?", 'fr': "Quel État americain est connu comme le 'Silver State' (État de l'argent) ?", 'da': "Hvilken amerikansk delstat er kendt som 'Silver State'?"},
 'opt1': {'en': 'Nevada', 'de': 'Nevada', 'es': 'Nevada', 'nl': 'Nevada', 'sv': 'Nevada', 'fr': 'le Nevada', 'da': 'Nevada'},
 'opt2': {'en': 'Colorado', 'de': 'Colorado', 'es': 'Colorado', 'nl': 'Colorado', 'sv': 'Colorado', 'fr': 'le Colorado', 'da': 'Colorado'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which orange root vegetable is used to make traditional carrot cake?', 'de': 'Welches orangefarbene Wurzelgemuese wird fuer traditionellen Karottenkuchen verwendet?', 'es': '¿Que verdura de raiz naranja se usa para hacer el pastel de zanahoria tradicional?', 'nl': 'Welke oranje wortelgroente wordt gebruikt om traditionele wortelcake te maken?', 'sv': 'Vilken orange rotfrukt anvands for att gora traditionell morotstarta?', 'fr': "Quel legume-racine orange est utilise pour faire le traditionnel gateau aux carottes ?", 'da': 'Hvilken orange rodfrugt bruges til at lave traditionel gulerodskage?'},
 'opt1': {'en': 'Sweet potato', 'de': 'Suesskartoffel', 'es': 'Camote', 'nl': 'Zoete aardappel', 'sv': 'Sotpotatis', 'fr': 'Patate douce', 'da': 'Sødkartoffel'},
 'opt2': {'en': 'Carrot', 'de': 'Karotte', 'es': 'Zanahoria', 'nl': 'Wortel', 'sv': 'Morot', 'fr': 'Carotte', 'da': 'Gulerod'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves hitting a small ball with a mallet through hoops on grass?', 'de': 'Bei welcher Sportart schlaegt man einen kleinen Ball mit einem Schlaeger durch Reifen auf Gras?', 'es': '¿Que deporte consiste en golpear una pequeña pelota con un mazo a traves de aros en el cesped?', 'nl': 'Bij welke sport sla je een kleine bal met een hamer door hoepels op gras?', 'sv': 'Vilken sport innebar att slå en liten boll med en klubba genom bagar pa grasmatta?', 'fr': "Quel sport consiste a frapper une petite balle avec un maillet a travers des arceaux sur l'herbe ?", 'da': 'Hvilken sport involverer at slå en lille bold med en kølle gennem buer på græs?'},
 'opt1': {'en': 'Croquet', 'de': 'Krocket', 'es': 'Croquet', 'nl': 'Croquet', 'sv': 'Krocket', 'fr': 'Croquet', 'da': 'Kroket'},
 'opt2': {'en': 'Golf', 'de': 'Golf', 'es': 'Golf', 'nl': 'Golf', 'sv': 'Golf', 'fr': 'Golf', 'da': 'Golf'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which large sea creature has eight arms and can change color to camouflage?', 'de': 'Welches grosse Meereswesen hat acht Arme und kann seine Farbe zur Tarnung aendern?', 'es': '¿Que gran criatura marina tiene ocho brazos y puede cambiar de color para camuflarse?', 'nl': 'Welk groot zeewezen heeft acht armen en kan van kleur veranderen om te camoufleren?', 'sv': 'Vilken stor havsvarelse har atta armar och kan byta farg for att kamouflera sig?', 'fr': "Quelle grande creature marine a huit bras et peut changer de couleur pour se camoufler ?", 'da': 'Hvilket stort havvæsen har otte arme og kan skifte farve for at camouflere sig?'},
 'opt1': {'en': 'Squid', 'de': 'Tintenfisch', 'es': 'Calamar', 'nl': 'Inktvis', 'sv': 'Bläckfisk (tioarmad)', 'fr': 'Calmar', 'da': 'Blæksprutte (tiarmet)'},
 'opt2': {'en': 'Octopus', 'de': 'Krake', 'es': 'Pulpo', 'nl': 'Octopus', 'sv': 'Bläckfisk', 'fr': 'Poulpe', 'da': 'Blæksprutte'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which holiday is celebrated with a large parade and fireworks on New Year's Eve?", 'de': 'Welcher Feiertag wird mit einer grossen Parade und Feuerwerk an Silvester gefeiert?', 'es': '¿Que festividad se celebra con un gran desfile y fuegos artificiales en Nochevieja?', 'nl': 'Welke feestdag wordt gevierd met een grote parade en vuurwerk op Oudjaarsavond?', 'sv': 'Vilken hogtid firas med en stor parad och fyrverkerier pa nyarsafton?', 'fr': "Quelle fete se celebre avec un grand defile et des feux d'artifice le soir du Nouvel An ?", 'da': 'Hvilken helligdag fejres med en stor parade og fyrværkeri på nytårsaften?'},
 'opt1': {'en': "New Year's Eve", 'de': 'Silvester', 'es': 'Nochevieja', 'nl': 'Oudjaarsavond', 'sv': 'Nyarsafton', 'fr': "La Saint-Sylvestre", 'da': 'Nytårsaften'},
 'opt2': {'en': 'Independence Day', 'de': 'Independence Day', 'es': 'Dia de la Independencia', 'nl': 'Independence Day', 'sv': 'Sjalvstandighetsdagen', 'fr': "Le jour de l'Independance", 'da': 'Uafhængighedsdag'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which continent is home to the Andes mountain range?', 'de': 'Auf welchem Kontinent liegt das Andengebirge?', 'es': '¿En que continente esta la cordillera de los Andes?', 'nl': 'Op welk continent ligt het Andesgebergte?', 'sv': 'Pa vilken kontinent ligger Anderna?', 'fr': "Sur quel continent se trouve la cordillere des Andes ?", 'da': 'Hvilket kontinent huser Andesbjergene?'},
 'opt1': {'en': 'North America', 'de': 'Nordamerika', 'es': 'Norteamerica', 'nl': 'Noord-Amerika', 'sv': 'Nordamerika', 'fr': "l'Amerique du Nord", 'da': 'Nordamerika'},
 'opt2': {'en': 'South America', 'de': 'Suedamerika', 'es': 'Sudamerica', 'nl': 'Zuid-Amerika', 'sv': 'Sydamerika', 'fr': "l'Amerique du Sud", 'da': 'Sydamerika'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which author wrote 'The Very Hungry Caterpillar'?", 'de': "Welcher Autor schrieb 'Die kleine Raupe Nimmersatt'?", 'es': "¿Que autor escribio 'La oruga muy hambrienta'?", 'nl': "Welke auteur schreef 'Rupsje Nooitgenoeg'?", 'sv': "Vilken forfattare skrev 'Den larvhungriga fjarilen'?", 'fr': "Quel auteur a ecrit 'La Chenille qui fait des trous' ?", 'da': "Hvilken forfatter skrev 'Den lille sultne larve'?"},
 'opt1': {'en': 'Maurice Sendak', 'de': 'Maurice Sendak', 'es': 'Maurice Sendak', 'nl': 'Maurice Sendak', 'sv': 'Maurice Sendak', 'fr': 'Maurice Sendak', 'da': 'Maurice Sendak'},
 'opt2': {'en': 'Eric Carle', 'de': 'Eric Carle', 'es': 'Eric Carle', 'nl': 'Eric Carle', 'sv': 'Eric Carle', 'fr': 'Eric Carle', 'da': 'Eric Carle'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic toy is a doll house filled with miniature furniture?', 'de': 'Welches klassische Spielzeug ist ein Puppenhaus voller Miniaturmoebel?', 'es': '¿Que juguete clasico es una casa de muñecas llena de muebles en miniatura?', 'nl': 'Welk klassiek speelgoed is een poppenhuis vol miniatuurmeubels?', 'sv': 'Vilken klassisk leksak ar ett dockhus fyllt med miniatyrmobler?', 'fr': "Quel jouet classique est une maison de poupees remplie de meubles miniatures ?", 'da': 'Hvilket klassisk legetøj er et dukkehus fyldt med miniaturemøbler?'},
 'opt1': {'en': 'Dollhouse', 'de': 'Puppenhaus', 'es': 'Casa de muñecas', 'nl': 'Poppenhuis', 'sv': 'Dockhus', 'fr': 'Maison de poupee', 'da': 'Dukkehus'},
 'opt2': {'en': 'Playset', 'de': 'Spielset', 'es': 'Set de juego', 'nl': 'Speelset', 'sv': 'Lekset', 'fr': 'Coffret de jeu', 'da': 'Legesæt'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which white spice is ground from peppercorns and used at nearly every meal?', 'de': 'Welches weisse Gewuerz wird aus Pfefferkoernern gemahlen und bei fast jeder Mahlzeit verwendet?', 'es': '¿Que especia blanca se muele de granos de pimienta y se usa en casi todas las comidas?', 'nl': 'Welke witte specerij wordt gemalen van peperkorrels en bij bijna elke maaltijd gebruikt?', 'sv': 'Vilken vit krydda mals fran pepparkorn och anvands vid nastan varje maltid?', 'fr': "Quelle epice blanche est moulue a partir de grains de poivre et utilisee a presque chaque repas ?", 'da': 'Hvilket hvidt krydderi males af peberkorn og bruges ved næsten hvert måltid?'},
 'opt1': {'en': 'White pepper', 'de': 'Weisser Pfeffer', 'es': 'Pimienta blanca', 'nl': 'Witte peper', 'sv': 'Vitpeppar', 'fr': 'Poivre blanc', 'da': 'Hvid peber'},
 'opt2': {'en': 'Black pepper', 'de': 'Schwarzer Pfeffer', 'es': 'Pimienta negra', 'nl': 'Zwarte peper', 'sv': 'Svartpeppar', 'fr': 'Poivre noir', 'da': 'Sort peber'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. holiday honors the birthday of a civil rights leader in January?', 'de': 'Welcher US-Feiertag im Januar ehrt den Geburtstag eines Buergerrechtlers?', 'es': '¿Que festividad de EE.UU. en enero honra el cumpleaños de un lider de derechos civiles?', 'nl': 'Welke Amerikaanse feestdag in januari eert de verjaardag van een burgerrechtenleider?', 'sv': 'Vilken amerikansk hogtid i januari hedrar fodelsedagen for en medborgarrattsledare?', 'fr': "Quelle fete americaine de janvier honore l'anniversaire d'un leader des droits civiques ?", 'da': 'Hvilken amerikansk helligdag i januar hædrer fødselsdagen for en borgerrettighedsleder?'},
 'opt1': {'en': 'Martin Luther King Jr. Day', 'de': 'Martin Luther King Jr. Day', 'es': 'Dia de Martin Luther King Jr.', 'nl': 'Martin Luther King Jr. Day', 'sv': 'Martin Luther King Jr.-dagen', 'fr': 'Le jour de Martin Luther King', 'da': 'Martin Luther King Jr. Day'},
 'opt2': {'en': "Presidents' Day", 'de': 'Presidents\' Day', 'es': 'Dia de los Presidentes', 'nl': 'Presidentsdag', 'sv': 'Presidentdagen', 'fr': 'Le jour des Presidents', 'da': 'Presidents’ Day'}},

{'category': 'Music', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which decade is associated with the rise of MTV and music videos?', 'de': 'Mit welchem Jahrzehnt wird der Aufstieg von MTV und Musikvideos assoziiert?', 'es': '¿Con que decada se asocia el auge de MTV y los videos musicales?', 'nl': 'Met welk decennium wordt de opkomst van MTV en muziekvideo\'s geassocieerd?', 'sv': 'Vilket artionde forknippas med framvaxten av MTV och musikvideor?', 'fr': "A quelle decennie associe-t-on la montee de MTV et des clips musicaux ?", 'da': 'Hvilket årti forbindes med fremkomsten af MTV og musikvideoer?'},
 'opt1': {'en': '1970s', 'de': '1970er', 'es': 'los años 70', 'nl': 'jaren 70', 'sv': '1970-talet', 'fr': 'les annees 1970', 'da': '1970’erne'},
 'opt2': {'en': '1980s', 'de': '1980er', 'es': 'los años 80', 'nl': 'jaren 80', 'sv': '1980-talet', 'fr': 'les annees 1980', 'da': '1980’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Rome?', 'de': 'In welchem Land liegt die antike Stadt Rom?', 'es': '¿En que pais esta la antigua ciudad de Roma?', 'nl': 'In welk land ligt de oude stad Rome?', 'sv': 'I vilket land ligger den antika staden Rom?', 'fr': "Dans quel pays se trouve l'ancienne cite de Rome ?", 'da': 'Hvilket land huser den antikke by Rom?'},
 'opt1': {'en': 'Italy', 'de': 'Italien', 'es': 'Italia', 'nl': 'Italie', 'sv': 'Italien', 'fr': "l'Italie", 'da': 'Italien'},
 'opt2': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which farm animal produces eggs used in baking?', 'de': 'Welches Nutztier produziert Eier, die zum Backen verwendet werden?', 'es': '¿Que animal de granja produce huevos usados para hornear?', 'nl': 'Welk boerderijdier produceert eieren die worden gebruikt bij het bakken?', 'sv': 'Vilket lantbruksdjur producerar agg som anvands vid bakning?', 'fr': "Quel animal de ferme produit des œufs utilises en patisserie ?", 'da': 'Hvilket landbrugsdyr producerer æg brugt til bagning?'},
 'opt1': {'en': 'Goose', 'de': 'Gans', 'es': 'Ganso', 'nl': 'Gans', 'sv': 'Gas', 'fr': 'Oie', 'da': 'Gås'},
 'opt2': {'en': 'Chicken', 'de': 'Huhn', 'es': 'Gallina', 'nl': 'Kip', 'sv': 'Höna', 'fr': 'Poule', 'da': 'Høne'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which gemstone is the birthstone for those born in August?', 'de': 'Welcher Edelstein ist der Geburtsstein fuer im August Geborene?', 'es': '¿Que piedra preciosa es la piedra de nacimiento para los nacidos en agosto?', 'nl': 'Welke edelsteen is de geboortesteen voor mensen geboren in augustus?', 'sv': 'Vilken adelsten ar fodelsestenen for de fodda i augusti?', 'fr': "Quelle pierre precieuse est la pierre de naissance d'aout ?", 'da': 'Hvilken ædelsten er fødselsstenen for dem, der er født i august?'},
 'opt1': {'en': 'Peridot', 'de': 'Peridot', 'es': 'Peridoto', 'nl': 'Peridoot', 'sv': 'Peridot', 'fr': 'Peridot', 'da': 'Peridot'},
 'opt2': {'en': 'Emerald', 'de': 'Smaragd', 'es': 'Esmeralda', 'nl': 'Smaragd', 'sv': 'Smaragd', 'fr': 'Emeraude', 'da': 'Smaragd'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic party game involves passing a parcel while music plays?', 'de': 'Bei welchem klassischen Partyspiel wird ein Paket weitergereicht, waehrend Musik spielt?', 'es': '¿Que juego de fiesta clasico consiste en pasar un paquete mientras suena musica?', 'nl': 'Bij welk klassiek feestspel geef je een pakketje door terwijl er muziek speelt?', 'sv': 'Vilken klassisk festlek innebar att skicka runt ett paket medan musik spelas?', 'fr': "Quel jeu de fete classique consiste a se passer un paquet pendant que la musique joue ?", 'da': 'Hvilken klassisk selskabsleg involverer at sende en pakke rundt, mens der spilles musik?'},
 'opt1': {'en': 'Musical Chairs', 'de': 'Reise nach Jerusalem', 'es': 'Sillas musicales', 'nl': 'Stoelendans', 'sv': 'Hela havet stormar', 'fr': 'Le jeu des chaises musicales', 'da': 'Musikalske stole'},
 'opt2': {'en': 'Pass the Parcel', 'de': 'Paket weiterreichen', 'es': 'Pasar el paquete', 'nl': 'Pakjes doorgeven', 'sv': 'Skicka paketet', 'fr': 'Le paquet cadeau (jeu)', 'da': 'Send pakken'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which red spice is dried and ground from sweet peppers, common in Spanish cooking?', 'de': 'Welches rote Gewuerz wird aus getrockneten und gemahlenen suessen Paprikaschoten hergestellt, ueblich in der spanischen Kueche?', 'es': '¿Que especia roja se seca y muele de pimientos dulces, comun en la cocina española?', 'nl': 'Welke rode specerij wordt gedroogd en gemalen van zoete paprika\'s, gebruikelijk in de Spaanse keuken?', 'sv': 'Vilken rod krydda torkas och mals fran soet paprika, vanlig i spansk mat?', 'fr': "Quelle epice rouge est sechee et moulue a partir de poivrons doux, courante dans la cuisine espagnole ?", 'da': 'Hvilket rødt krydderi er tørret og malet af søde peberfrugter, almindeligt i spansk mad?'},
 'opt1': {'en': 'Paprika', 'de': 'Paprika', 'es': 'Pimenton', 'nl': 'Paprikapoeder', 'sv': 'Paprikapulver', 'fr': 'Le paprika', 'da': 'Paprika'},
 'opt2': {'en': 'Chili powder', 'de': 'Chilipulver', 'es': 'Chile en polvo', 'nl': 'Chilipoeder', 'sv': 'Chilipulver', 'fr': 'La poudre de chili', 'da': 'Chilipulver'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which flower is traditionally worn by brides in a bouquet?', 'de': 'Welche Blume wird traditionell von Bräuten im Brautstrauss getragen?', 'es': '¿Que flor llevan tradicionalmente las novias en el ramo?', 'nl': 'Welke bloem wordt traditioneel door bruiden gedragen in een boeket?', 'sv': 'Vilken blomma baers traditionellt av brudar i en bukett?', 'fr': "Quelle fleur les mariees portent-elles traditionnellement dans un bouquet ?", 'da': 'Hvilken blomst bæres traditionelt af brude i en buket?'},
 'opt1': {'en': 'Lily', 'de': 'Lilie', 'es': 'Lirio', 'nl': 'Lelie', 'sv': 'Lilja', 'fr': 'Lys', 'da': 'Lilje'},
 'opt2': {'en': 'Rose', 'de': 'Rose', 'es': 'Rosa', 'nl': 'Roos', 'sv': 'Ros', 'fr': 'Rose', 'da': 'Rose'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Kennedy Space Center?', 'de': 'In welchem US-Bundesstaat liegt das Kennedy Space Center?', 'es': '¿En que estado de EE.UU. esta el Centro Espacial Kennedy?', 'nl': 'In welke Amerikaanse staat ligt het Kennedy Space Center?', 'sv': 'I vilken amerikansk delstat ligger Kennedy Space Center?', 'fr': "Dans quel État americain se trouve le Kennedy Space Center ?", 'da': 'I hvilken amerikansk delstat ligger Kennedy Space Center?'},
 'opt1': {'en': 'Florida', 'de': 'Florida', 'es': 'Florida', 'nl': 'Florida', 'sv': 'Florida', 'fr': 'la Floride', 'da': 'Florida'},
 'opt2': {'en': 'Texas', 'de': 'Texas', 'es': 'Texas', 'nl': 'Texas', 'sv': 'Texas', 'fr': 'le Texas', 'da': 'Texas'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which craft involves creating a picture using small colored tiles or stones?', 'de': 'Bei welchem Handwerk wird ein Bild aus kleinen farbigen Fliesen oder Steinen erstellt?', 'es': '¿Que manualidad implica crear una imagen usando pequeños azulejos o piedras de colores?', 'nl': 'Bij welk handwerk creëer je een afbeelding met kleine gekleurde tegels of steentjes?', 'sv': 'Vilket hantverk innebar att skapa en bild med sma fargade kakel eller stenar?', 'fr': "Quel art manuel consiste a creer une image avec de petits carreaux ou pierres colores ?", 'da': 'Hvilket håndværk involverer at skabe et billede med små farvede fliser eller sten?'},
 'opt1': {'en': 'Stained glass', 'de': 'Buntglas', 'es': 'Vidriera', 'nl': 'Glas-in-lood', 'sv': 'Glasmalning', 'fr': 'Le vitrail', 'da': 'Glasmaleri'},
 'opt2': {'en': 'Mosaic', 'de': 'Mosaik', 'es': 'Mosaico', 'nl': 'Mozaiek', 'sv': 'Mosaik', 'fr': 'La mosaique', 'da': 'Mosaik'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves jumping over a horizontal bar using a pole?', 'de': 'Bei welcher Sportart springt man mit einer Stange ueber eine waagerechte Latte?', 'es': '¿Que deporte consiste en saltar sobre una barra horizontal usando una pertiga?', 'nl': 'Bij welke sport spring je over een horizontale lat met behulp van een polsstok?', 'sv': 'Vilken sport innebar att hoppa over en horisontell ribba med en stav?', 'fr': "Quel sport consiste a sauter par-dessus une barre horizontale a l'aide d'une perche ?", 'da': 'Hvilken sport involverer at hoppe over en vandret bom ved hjælp af en stav?'},
 'opt1': {'en': 'Pole vault', 'de': 'Stabhochsprung', 'es': 'Salto con pertiga', 'nl': 'Polsstokhoogspringen', 'sv': 'Stavhopp', 'fr': 'Le saut a la perche', 'da': 'Stangspring'},
 'opt2': {'en': 'High jump', 'de': 'Hochsprung', 'es': 'Salto de altura', 'nl': 'Hoogspringen', 'sv': 'Hojdhopp', 'fr': 'Le saut en hauteur', 'da': 'Højdespring'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which item of clothing is worn on the hands to keep them warm in winter?', 'de': 'Welches Kleidungsstueck wird an den Haenden getragen, um sie im Winter warm zu halten?', 'es': '¿Que prenda se lleva en las manos para mantenerlas calientes en invierno?', 'nl': 'Welk kledingstuk wordt aan de handen gedragen om ze warm te houden in de winter?', 'sv': 'Vilket klädesplagg baers pa handerna for att halla dem varma pa vintern?', 'fr': "Quel vetement se porte aux mains pour les tenir au chaud en hiver ?", 'da': 'Hvilken beklædningsgenstand bæres på hænderne for at holde dem varme om vinteren?'},
 'opt1': {'en': 'Mittens', 'de': 'Fausthandschuhe', 'es': 'Manoplas', 'nl': 'Wanten', 'sv': 'Vantar', 'fr': 'Moufles', 'da': 'Vanter'},
 'opt2': {'en': 'Gloves', 'de': 'Handschuhe', 'es': 'Guantes', 'nl': 'Handschoenen', 'sv': 'Handskar', 'fr': 'Gants', 'da': 'Handsker'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic cartoon features a cat and mouse who constantly chase each other?', 'de': 'In welchem klassischen Cartoon jagen sich eine Katze und eine Maus staendig gegenseitig?', 'es': '¿Que dibujo animado clasico presenta a un gato y un raton que se persiguen constantemente?', 'nl': 'In welke klassieke tekenfilm achtervolgen een kat en muis elkaar voortdurend?', 'sv': 'I vilken klassisk tecknad film jagar en katt och en mus varandra konstant?', 'fr': "Quel dessin anime classique met en scene un chat et une souris qui se pourchassent constamment ?", 'da': 'Hvilken klassisk tegnefilm har en kat og mus, der konstant jagter hinanden?'},
 'opt1': {'en': 'Tom and Jerry', 'de': 'Tom und Jerry', 'es': 'Tom y Jerry', 'nl': 'Tom en Jerry', 'sv': 'Tom och Jerry', 'fr': 'Tom et Jerry', 'da': 'Tom og Jerry'},
 'opt2': {'en': 'Itchy and Scratchy', 'de': 'Itchy und Scratchy', 'es': 'Rasca y Pica', 'nl': 'Itchy en Scratchy', 'sv': 'Itchy och Scratchy', 'fr': 'Itchy et Scratchy', 'da': 'Itchy og Scratchy'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Ephesus and the ruins of Troy?', 'de': 'In welchem Land liegen die antike Stadt Ephesus und die Ruinen von Troja?', 'es': '¿En que pais estan la antigua ciudad de Efeso y las ruinas de Troya?', 'nl': 'In welk land liggen de oude stad Efeze en de ruines van Troje?', 'sv': 'I vilket land ligger den antika staden Efesos och ruinerna av Troja?', 'fr': "Dans quel pays se trouvent l'ancienne cite d'Ephese et les ruines de Troie ?", 'da': 'Hvilket land huser den antikke by Efesos og ruinerne af Troja?'},
 'opt1': {'en': 'Turkey', 'de': 'Tuerkei', 'es': 'Turquia', 'nl': 'Turkije', 'sv': 'Turkiet', 'fr': 'la Turquie', 'da': 'Tyrkiet'},
 'opt2': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was known for launching a campaign against childhood obesity?', 'de': 'Welche First Lady war bekannt fuer den Start einer Kampagne gegen Fettleibigkeit bei Kindern?', 'es': '¿Que primera dama fue conocida por lanzar una campaña contra la obesidad infantil?', 'nl': 'Welke First Lady stond bekend om het lanceren van een campagne tegen obesitas bij kinderen?', 'sv': 'Vilken First Lady var kand for att lansera en kampanj mot barnfetma?', 'fr': "Quelle Première dame etait connue pour avoir lance une campagne contre l'obesite infantile ?", 'da': 'Hvilken førstedame var kendt for at lancere en kampagne mod børnefedme?'},
 'opt1': {'en': 'Michelle Obama', 'de': 'Michelle Obama', 'es': 'Michelle Obama', 'nl': 'Michelle Obama', 'sv': 'Michelle Obama', 'fr': 'Michelle Obama', 'da': 'Michelle Obama'},
 'opt2': {'en': 'Laura Bush', 'de': 'Laura Bush', 'es': 'Laura Bush', 'nl': 'Laura Bush', 'sv': 'Laura Bush', 'fr': 'Laura Bush', 'da': 'Laura Bush'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Chinese dish is stir-fried rice with vegetables, egg, and often meat?', 'de': 'Welches chinesische Gericht ist gebratener Reis mit Gemuese, Ei und oft Fleisch?', 'es': '¿Que plato chino es arroz salteado con verduras, huevo y a menudo carne?', 'nl': 'Welk Chinees gerecht is gebakken rijst met groenten, ei en vaak vlees?', 'sv': 'Vilken kinesisk ratt ar wokat ris med gronsaker, agg och ofta kott?', 'fr': "Quel plat chinois est du riz saute avec des legumes, un oeuf et souvent de la viande ?", 'da': 'Hvilken kinesisk ret er stegt ris med grøntsager, æg og ofte kød?'},
 'opt1': {'en': 'Lo mein', 'de': 'Lo Mein', 'es': 'Lo mein', 'nl': 'Lo mein', 'sv': 'Lo mein', 'fr': 'Lo mein', 'da': 'Lo mein'},
 'opt2': {'en': 'Fried rice', 'de': 'Gebratener Reis', 'es': 'Arroz frito', 'nl': 'Gebakken rijst', 'sv': 'Wokat ris', 'fr': 'Riz saute', 'da': 'Stegt ris'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which body part produces sperm and testosterone?', 'de': 'Welcher Koerperteil produziert Spermien und Testosteron?', 'es': '¿Que parte del cuerpo produce esperma y testosterona?', 'nl': 'Welk lichaamsdeel produceert sperma en testosteron?', 'sv': 'Vilken kroppsdel producerar spermier och testosteron?', 'fr': "Quelle partie du corps produit le sperme et la testosterone ?", 'da': 'Hvilken kropsdel producerer sæd og testosteron?'},
 'opt1': {'en': 'Prostate', 'de': 'Prostata', 'es': 'Prostata', 'nl': 'Prostaat', 'sv': 'Prostata', 'fr': 'Prostate', 'da': 'Prostata'},
 'opt2': {'en': 'Testes', 'de': 'Hoden', 'es': 'Testiculos', 'nl': 'Testikels', 'sv': 'Testiklar', 'fr': 'Testicules', 'da': 'Testikler'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'The Catcher in the Rye'?", 'de': "Welcher Autor schrieb 'Der Faenger im Roggen'?", 'es': "¿Que autor escribio 'El guardian entre el centeno'?", 'nl': "Welke auteur schreef 'De vanger in het graan'?", 'sv': "Vilken forfattare skrev 'Ratten i vardagsrummet'?", 'fr': "Quel auteur a ecrit 'L'Attrape-coeurs' ?", 'da': "Hvilken forfatter skrev 'Griberen i rugen'?"},
 'opt1': {'en': 'John Steinbeck', 'de': 'John Steinbeck', 'es': 'John Steinbeck', 'nl': 'John Steinbeck', 'sv': 'John Steinbeck', 'fr': 'John Steinbeck', 'da': 'John Steinbeck'},
 'opt2': {'en': 'J.D. Salinger', 'de': 'J.D. Salinger', 'es': 'J.D. Salinger', 'nl': 'J.D. Salinger', 'sv': 'J.D. Salinger', 'fr': 'J.D. Salinger', 'da': 'J.D. Salinger'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Peggy Olson in the TV series 'Mad Men'?", 'de': "Welche Schauspielerin spielte Peggy Olson in der TV-Serie 'Mad Men'?", 'es': "¿Que actriz interpreto a Peggy Olson en la serie de TV 'Mad Men'?", 'nl': "Welke actrice speelde Peggy Olson in de tv-serie 'Mad Men'?", 'sv': "Vilken skadespelerska spelade Peggy Olson i TV-serien 'Mad Men'?", 'fr': "Quelle actrice a joue Peggy Olson dans la serie 'Mad Men' ?", 'da': "Hvilken skuespiller spillede Peggy Olson i tv-serien 'Mad Men'?"},
 'opt1': {'en': 'Elisabeth Moss', 'de': 'Elisabeth Moss', 'es': 'Elisabeth Moss', 'nl': 'Elisabeth Moss', 'sv': 'Elisabeth Moss', 'fr': 'Elisabeth Moss', 'da': 'Elisabeth Moss'},
 'opt2': {'en': 'Christina Hendricks', 'de': 'Christina Hendricks', 'es': 'Christina Hendricks', 'nl': 'Christina Hendricks', 'sv': 'Christina Hendricks', 'fr': 'Christina Hendricks', 'da': 'Christina Hendricks'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which singer is known as the 'Material Girl'?", 'de': "Welche Saengerin wird als 'Material Girl' bezeichnet?", 'es': "¿Que cantante es conocida como la 'Material Girl'?", 'nl': "Welke zangeres staat bekend als de 'Material Girl'?", 'sv': "Vilken sangerska kallas 'Material Girl'?", 'fr': "Quelle chanteuse est surnommee 'Material Girl' ?", 'da': "Hvilken sanger er kendt som 'Material Girl'?"},
 'opt1': {'en': 'Cyndi Lauper', 'de': 'Cyndi Lauper', 'es': 'Cyndi Lauper', 'nl': 'Cyndi Lauper', 'sv': 'Cyndi Lauper', 'fr': 'Cyndi Lauper', 'da': 'Cyndi Lauper'},
 'opt2': {'en': 'Madonna', 'de': 'Madonna', 'es': 'Madonna', 'nl': 'Madonna', 'sv': 'Madonna', 'fr': 'Madonna', 'da': 'Madonna'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Chichen Itza?', 'de': 'In welchem Land liegt die antike Stadt Chichen Itza?', 'es': '¿En que pais esta la antigua ciudad de Chichen Itza?', 'nl': 'In welk land ligt de oude stad Chichen Itza?', 'sv': 'I vilket land ligger den antika staden Chichen Itza?', 'fr': "Dans quel pays se trouve l'ancienne cite de Chichen Itza ?", 'da': 'Hvilket land huser den antikke by Chichen Itza?'},
 'opt1': {'en': 'Belize', 'de': 'Belize', 'es': 'Belice', 'nl': 'Belize', 'sv': 'Belize', 'fr': 'le Belize', 'da': 'Belize'},
 'opt2': {'en': 'Mexico', 'de': 'Mexiko', 'es': 'Mexico', 'nl': 'Mexico', 'sv': 'Mexiko', 'fr': 'le Mexique', 'da': 'Mexico'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which large land mammal has a single horn on its nose, found in Africa and Asia?', 'de': 'Welches grosse Landsaeugetier hat ein einzelnes Horn auf der Nase und lebt in Afrika und Asien?', 'es': '¿Que gran mamifero terrestre tiene un unico cuerno en la nariz, encontrado en Africa y Asia?', 'nl': 'Welk groot landzoogdier heeft een enkele hoorn op zijn neus, gevonden in Afrika en Azie?', 'sv': 'Vilket stort landdaggdjur har ett enda horn pa nasan, finns i Afrika och Asien?', 'fr': "Quel grand mammifere terrestre a une seule corne sur le nez, present en Afrique et en Asie ?", 'da': 'Hvilket stort landpattedyr har et enkelt horn på næsen, fundet i Afrika og Asien?'},
 'opt1': {'en': 'Hippopotamus', 'de': 'Nilpferd', 'es': 'Hipopotamo', 'nl': 'Nijlpaard', 'sv': 'Flodhast', 'fr': 'Hippopotame', 'da': 'Flodhest'},
 'opt2': {'en': 'Rhinoceros', 'de': 'Nashorn', 'es': 'Rinoceronte', 'nl': 'Neushoorn', 'sv': 'Noshorning', 'fr': 'Rhinoceros', 'da': 'Næsehorn'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Vietnamese sandwich combines French baguette with pickled vegetables and meat?', 'de': 'Welches vietnamesische Sandwich kombiniert franzoesisches Baguette mit eingelegtem Gemuese und Fleisch?', 'es': '¿Que sandwich vietnamita combina baguette francesa con verduras encurtidas y carne?', 'nl': 'Welke Vietnamese sandwich combineert Frans stokbrood met ingemaakte groenten en vlees?', 'sv': 'Vilken vietnamesisk smorgas kombinerar fransk baguette med inlagda gronsaker och kott?', 'fr': "Quel sandwich vietnamien combine baguette française avec des legumes marines et de la viande ?", 'da': 'Hvilken vietnamesisk sandwich kombinerer fransk baguette med syltede grøntsager og kød?'},
 'opt1': {'en': 'Banh mi', 'de': 'Banh Mi', 'es': 'Banh mi', 'nl': 'Banh mi', 'sv': 'Banh mi', 'fr': 'Banh mi', 'da': 'Banh mi'},
 'opt2': {'en': 'Spring roll', 'de': 'Fruehlingsrolle', 'es': 'Rollito de primavera', 'nl': 'Loempia', 'sv': 'Vaarrulle', 'fr': 'Rouleau de printemps', 'da': 'Forårsrulle'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which First Lady was the first to be photographed extensively for magazines and fashion coverage?', 'de': 'Welche First Lady wurde als erste ausgiebig fuer Zeitschriften und Modeberichterstattung fotografiert?', 'es': '¿Que primera dama fue la primera en ser fotografiada extensamente para revistas y cobertura de moda?', 'nl': 'Welke First Lady werd als eerste uitgebreid gefotografeerd voor tijdschriften en modeverslaggeving?', 'sv': 'Vilken First Lady var forst med att fotograferas extensivt for tidningar och modebevakning?', 'fr': "Quelle Première dame a ete la premiere a etre abondamment photographiee pour des magazines et de la couverture mode ?", 'da': 'Hvilken førstedame var den første til at blive fotograferet omfattende til blade og modedækning?'},
 'opt1': {'en': 'Mamie Eisenhower', 'de': 'Mamie Eisenhower', 'es': 'Mamie Eisenhower', 'nl': 'Mamie Eisenhower', 'sv': 'Mamie Eisenhower', 'fr': 'Mamie Eisenhower', 'da': 'Mamie Eisenhower'},
 'opt2': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which artist is known for the painting 'Girl with a Pearl Earring'?", 'de': "Welcher Kuenstler ist fuer das Gemaelde 'Das Maedchen mit dem Perlenohrring' bekannt?", 'es': "¿Que artista es conocido por la pintura 'La joven de la perla'?", 'nl': "Welke kunstenaar staat bekend om het schilderij 'Meisje met de parel'?", 'sv': "Vilken konstnar ar kand for malningen 'Flicka med parlorhange'?", 'fr': "Quel artiste est connu pour le tableau 'La Jeune Fille a la perle' ?", 'da': "Hvilken kunstner er kendt for maleriet 'Pigen med perleøreringen'?"},
 'opt1': {'en': 'Johannes Vermeer', 'de': 'Johannes Vermeer', 'es': 'Johannes Vermeer', 'nl': 'Johannes Vermeer', 'sv': 'Johannes Vermeer', 'fr': 'Johannes Vermeer', 'da': 'Johannes Vermeer'},
 'opt2': {'en': 'Jan Steen', 'de': 'Jan Steen', 'es': 'Jan Steen', 'nl': 'Jan Steen', 'sv': 'Jan Steen', 'fr': 'Jan Steen', 'da': 'Jan Steen'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which flowering shrub is known for its clusters of small flowers and comes in blue, pink, or purple depending on soil acidity?', 'de': 'Welcher bluehende Strauch ist fuer seine Buescheln kleiner Blueten bekannt und kommt je nach Bodensaeure in Blau, Rosa oder Lila vor?', 'es': '¿Que arbusto en flor es conocido por sus racimos de pequeñas flores y viene en azul, rosa o morado segun la acidez del suelo?', 'nl': 'Welke bloeiende struik staat bekend om zijn trosjes kleine bloemen en komt voor in blauw, roze of paars afhankelijk van de bodemzuurgraad?', 'sv': 'Vilken blommande buske ar kand for sina klasar av sma blommor och kommer i blatt, rosa eller lila beroende pa jordens surhetsgrad?', 'fr': "Quel arbuste a fleurs est connu pour ses grappes de petites fleurs et se decline en bleu, rose ou violet selon l'acidite du sol ?", 'da': 'Hvilken blomstrende busk er kendt for sine klynger af små blomster og fås i blå, pink eller lilla afhængigt af jordens surhedsgrad?'},
 'opt1': {'en': 'Rhododendron', 'de': 'Rhododendron', 'es': 'Rododendro', 'nl': 'Rododendron', 'sv': 'Rhododendron', 'fr': 'Rhododendron', 'da': 'Rhododendron'},
 'opt2': {'en': 'Hydrangea', 'de': 'Hortensie', 'es': 'Hortensia', 'nl': 'Hortensia', 'sv': 'Hortensia', 'fr': 'Hortensia', 'da': 'Hortensia'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which musical features the song 'Defying Gravity' from a story about the witches of Oz?", 'de': "In welchem Musical kommt das Lied 'Defying Gravity' vor, aus einer Geschichte ueber die Hexen von Oz?", 'es': "¿Que musical presenta la cancion 'Defying Gravity' de una historia sobre las brujas de Oz?", 'nl': "Welke musical bevat het lied 'Defying Gravity' uit een verhaal over de heksen van Oz?", 'sv': "Vilken musikal innehaller latten 'Defying Gravity' fran en historia om Oz haxor?", 'fr': "Quelle comedie musicale comporte la chanson 'Defying Gravity' tiree d'une histoire sur les sorcieres d'Oz ?", 'da': "Hvilken musical har sangen 'Defying Gravity' fra en historie om heksene i Oz?"},
 'opt1': {'en': 'The Wiz', 'de': 'The Wiz', 'es': 'The Wiz', 'nl': 'The Wiz', 'sv': 'The Wiz', 'fr': 'The Wiz', 'da': 'The Wiz'},
 'opt2': {'en': 'Wicked', 'de': 'Wicked', 'es': 'Wicked', 'nl': 'Wicked', 'sv': 'Wicked', 'fr': 'Wicked', 'da': 'Wicked'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country has Bern as its capital?', 'de': 'Welches Land hat Bern als Hauptstadt?', 'es': '¿Que pais tiene a Berna como su capital?', 'nl': 'Welk land heeft Bern als hoofdstad?', 'sv': 'Vilket land har Bern som huvudstad?', 'fr': "Quel pays a Berne pour capitale ?", 'da': 'Hvilket land har Bern som hovedstad?'},
 'opt1': {'en': 'Austria', 'de': 'Österreich', 'es': 'Austria', 'nl': 'Oostenrijk', 'sv': 'Osterrike', 'fr': "l'Autriche", 'da': 'Østrig'},
 'opt2': {'en': 'Switzerland', 'de': 'Schweiz', 'es': 'Suiza', 'nl': 'Zwitserland', 'sv': 'Schweiz', 'fr': 'la Suisse', 'da': 'Schweiz'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which spice is used in traditional apple pie alongside cinnamon and nutmeg?', 'de': 'Welches Gewuerz wird in traditionellem Apfelkuchen neben Zimt und Muskatnuss verwendet?', 'es': '¿Que especia se usa en el pastel de manzana tradicional junto a la canela y la nuez moscada?', 'nl': 'Welke specerij wordt gebruikt in traditionele appeltaart naast kaneel en nootmuskaat?', 'sv': 'Vilken krydda anvands i traditionell appelpaj tillsammans med kanel och muskot?', 'fr': "Quelle epice est utilisee dans la tarte aux pommes traditionnelle aux cotes de la cannelle et de la muscade ?", 'da': 'Hvilket krydderi bruges i traditionel æbletærte sammen med kanel og muskatnød?'},
 'opt1': {'en': 'Cardamom', 'de': 'Kardamom', 'es': 'Cardamomo', 'nl': 'Kardemom', 'sv': 'Kardemumma', 'fr': 'La cardamome', 'da': 'Kardemomme'},
 'opt2': {'en': 'Allspice', 'de': 'Piment', 'es': 'Pimienta de Jamaica', 'nl': 'Piment', 'sv': 'Kryddpeppar', 'fr': 'Le quatre-epices', 'da': 'Allehånde'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which craft involves creating patterns by folding and dyeing fabric, popular in Japan?', 'de': 'Bei welchem Handwerk werden Muster durch Falten und Faerben von Stoff erzeugt, beliebt in Japan?', 'es': '¿Que manualidad implica crear patrones doblando y tiñendo tela, popular en Japon?', 'nl': 'Bij welk handwerk creëer je patronen door stof te vouwen en te verven, populair in Japan?', 'sv': 'Vilket hantverk innebar att skapa monster genom att vika och farga tyg, populart i Japan?', 'fr': "Quel art manuel consiste a creer des motifs en pliant et teignant du tissu, populaire au Japon ?", 'da': 'Hvilket håndværk involverer at skabe mønstre ved at folde og farve stof, populært i Japan?'},
 'opt1': {'en': 'Shibori', 'de': 'Shibori', 'es': 'Shibori', 'nl': 'Shibori', 'sv': 'Shibori', 'fr': 'Shibori', 'da': 'Shibori'},
 'opt2': {'en': 'Batik', 'de': 'Batik', 'es': 'Batik', 'nl': 'Batik', 'sv': 'Batik', 'fr': 'Le batik', 'da': 'Batik'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which part of the eye focuses light onto the retina?', 'de': 'Welcher Teil des Auges buendelt Licht auf die Netzhaut?', 'es': '¿Que parte del ojo enfoca la luz sobre la retina?', 'nl': 'Welk deel van het oog richt licht op het netvlies?', 'sv': 'Vilken del av ogat fokuserar ljus pa natthinnan?', 'fr': "Quelle partie de l'oeil focalise la lumiere sur la retine ?", 'da': 'Hvilken del af øjet fokuserer lys på nethinden?'},
 'opt1': {'en': 'Lens', 'de': 'Linse', 'es': 'Cristalino', 'nl': 'Ooglens', 'sv': 'Lins', 'fr': 'Cristallin', 'da': 'Linse'},
 'opt2': {'en': 'Cornea', 'de': 'Hornhaut', 'es': 'Cornea', 'nl': 'Hoornvlies', 'sv': 'Hornhinnan', 'fr': 'Cornee', 'da': 'Hornhinde'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'In soccer, what is the term for a shot from a stationary ball awarded for a foul in the box?', 'de': 'Wie nennt man beim Fussball einen Schuss von einem ruhenden Ball, der fuer ein Foul im Strafraum verhaengt wird?', 'es': '¿En futbol, como se llama un tiro desde una pelota parada otorgado por una falta en el area?', 'nl': 'Hoe heet bij voetbal een schot van een stilliggende bal toegekend voor een overtreding in het strafschopgebied?', 'sv': 'Vad kallas i fotboll ett skott fran en stillastaende boll som ges for en foul i straffomradet?', 'fr': "En football, comment appelle-t-on un tir sur balle arretee accorde pour une faute dans la surface ?", 'da': 'Hvad kaldes i fodbold et skud fra en stillestående bold, tildelt for en forseelse i feltet?'},
 'opt1': {'en': 'Free kick', 'de': 'Freistoss', 'es': 'Tiro libre', 'nl': 'Vrije trap', 'sv': 'Frispark', 'fr': 'Le coup franc', 'da': 'Frispark'},
 'opt2': {'en': 'Penalty kick', 'de': 'Elfmeter', 'es': 'Penalti', 'nl': 'Strafschop', 'sv': 'Straffspark', 'fr': 'Le penalty', 'da': 'Straffespark'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Greek region is famous for its ancient Olympic Games origin?', 'de': 'Welche griechische Region ist beruehmt fuer den antiken Ursprung der Olympischen Spiele?', 'es': '¿Que region griega es famosa por ser el origen de los antiguos Juegos Olimpicos?', 'nl': 'Welke Griekse regio staat bekend als de oorsprong van de antieke Olympische Spelen?', 'sv': 'Vilken grekisk region ar kand som ursprunget till de antika olympiska spelen?', 'fr': "Quelle region grecque est celebre pour etre l'origine des Jeux olympiques antiques ?", 'da': 'Hvilken græsk region er berømt som oprindelsen til de antikke olympiske lege?'},
 'opt1': {'en': 'Olympia', 'de': 'Olympia', 'es': 'Olimpia', 'nl': 'Olympia', 'sv': 'Olympia', 'fr': 'Olympie', 'da': 'Olympia'},
 'opt2': {'en': 'Delphi', 'de': 'Delphi', 'es': 'Delfos', 'nl': 'Delphi', 'sv': 'Delfi', 'fr': 'Delphes', 'da': 'Delfi'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Sally Albright in 'When Harry Met Sally'?", 'de': "Welche Schauspielerin spielte Sally Albright in 'Harry und Sally'?", 'es': "¿Que actriz interpreto a Sally Albright en 'Cuando Harry encontro a Sally'?", 'nl': "Welke actrice speelde Sally Albright in 'When Harry Met Sally'?", 'sv': "Vilken skadespelerska spelade Sally Albright i 'Harry och Sally'?", 'fr': "Quelle actrice a joue Sally Albright dans 'Quand Harry rencontre Sally' ?", 'da': "Hvilken skuespiller spillede Sally Albright i 'When Harry Met Sally'?"},
 'opt1': {'en': 'Meg Ryan', 'de': 'Meg Ryan', 'es': 'Meg Ryan', 'nl': 'Meg Ryan', 'sv': 'Meg Ryan', 'fr': 'Meg Ryan', 'da': 'Meg Ryan'},
 'opt2': {'en': 'Julia Roberts', 'de': 'Julia Roberts', 'es': 'Julia Roberts', 'nl': 'Julia Roberts', 'sv': 'Julia Roberts', 'fr': 'Julia Roberts', 'da': 'Julia Roberts'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which classic novel features the character Scarlett O'Hara?", 'de': "In welchem Klassiker kommt die Figur Scarlett O'Hara vor?", 'es': "¿Que novela clasica presenta al personaje Scarlett O'Hara?", 'nl': "In welke klassieke roman komt het personage Scarlett O'Hara voor?", 'sv': "I vilken klassisk roman forekommer karaktaren Scarlett O'Hara?", 'fr': "Quel roman classique met en scene le personnage de Scarlett O'Hara ?", 'da': "Hvilken klassisk roman har karakteren Scarlett O'Hara?"},
 'opt1': {'en': 'Gone with the Wind', 'de': 'Vom Winde verweht', 'es': 'Lo que el viento se llevo', 'nl': 'Gejaagd door de Wind', 'sv': 'Borta med vinden', 'fr': 'Autant en emporte le vent', 'da': 'Borte med blæsten'},
 'opt2': {'en': 'The Help', 'de': 'The Help', 'es': 'Criadas y señoras', 'nl': 'The Help', 'sv': 'Niceville', 'fr': 'La Couleur des sentiments', 'da': 'The Help'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which 1970s band is known for 'Hotel California'?", 'de': "Welche Band der 1970er ist fuer 'Hotel California' bekannt?", 'es': "¿Que banda de los 70 es conocida por 'Hotel California'?", 'nl': "Welke band uit de jaren 70 staat bekend om 'Hotel California'?", 'sv': "Vilket 1970-talsband ar kant for 'Hotel California'?", 'fr': "Quel groupe des annees 1970 est connu pour 'Hotel California' ?", 'da': "Hvilket 1970’er-band er kendt for 'Hotel California'?"},
 'opt1': {'en': 'The Eagles', 'de': 'The Eagles', 'es': 'The Eagles', 'nl': 'The Eagles', 'sv': 'The Eagles', 'fr': 'The Eagles', 'da': 'The Eagles'},
 'opt2': {'en': 'Fleetwood Mac', 'de': 'Fleetwood Mac', 'es': 'Fleetwood Mac', 'nl': 'Fleetwood Mac', 'sv': 'Fleetwood Mac', 'fr': 'Fleetwood Mac', 'da': 'Fleetwood Mac'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a rich, buttery biscuit often served with gravy?', 'de': 'Welches suddstaatliche Gericht ist ein reichhaltiges, butteriges Biscuit, oft mit Sauce serviert?', 'es': '¿Que plato sureño de EE.UU. es un biscuit rico y mantecoso a menudo servido con salsa?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een rijk, boterachtig biscuit vaak geserveerd met jus?', 'sv': 'Vilken sydstatlig ratt ar en rik, smorig biscuit ofta serverad med sas?', 'fr': "Quel plat du Sud des États-Unis est un biscuit riche et beurre souvent servi avec une sauce ?", 'da': 'Hvilken sydstatsret er en rig, smøragtig biscuit ofte serveret med sovs?'},
 'opt1': {'en': 'Biscuits and gravy', 'de': 'Biscuits and Gravy', 'es': 'Biscuits con salsa', 'nl': 'Biscuits met jus', 'sv': 'Biscuits med sas', 'fr': 'Biscuits a la sauce', 'da': 'Biscuits med sovs'},
 'opt2': {'en': 'Cornbread', 'de': 'Maisbrot', 'es': 'Pan de maiz', 'nl': 'Maisbrood', 'sv': 'Majsbröd', 'fr': 'Pain de mais', 'da': 'Majsbrød'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which queen ruled England during the defeat of the Spanish Armada?', 'de': 'Welche Koenigin regierte England waehrend der Niederlage der Spanischen Armada?', 'es': '¿Que reina goberno Inglaterra durante la derrota de la Armada Española?', 'nl': 'Welke koningin regeerde Engeland tijdens de nederlaag van de Spaanse Armada?', 'sv': 'Vilken drottning styrde England under nederlaget for spanska armadan?', 'fr': "Quelle reine a regne sur l'Angleterre lors de la defaite de l'Invincible Armada espagnole ?", 'da': 'Hvilken dronning regerede England under nederlaget for den spanske armada?'},
 'opt1': {'en': 'Mary I', 'de': 'Maria I.', 'es': 'Maria I', 'nl': 'Mary I', 'sv': 'Maria I', 'fr': 'Marie Ire', 'da': 'Maria 1.'},
 'opt2': {'en': 'Elizabeth I', 'de': 'Elisabeth I.', 'es': 'Isabel I', 'nl': 'Elizabeth I', 'sv': 'Elizabeth I', 'fr': 'Elisabeth Ire', 'da': 'Elizabeth 1.'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large aquatic mammal is known for its intelligence and is trained at marine parks?', 'de': 'Welches grosse Wassersaeugetier ist fuer seine Intelligenz bekannt und wird in Meeresparks trainiert?', 'es': '¿Que gran mamifero acuatico es conocido por su inteligencia y se entrena en parques marinos?', 'nl': 'Welk groot waterzoogdier staat bekend om zijn intelligentie en wordt getraind in zeeparken?', 'sv': 'Vilket stort vattenlevande daggdjur ar kant for sin intelligens och trans i marina parker?', 'fr': "Quel grand mammifere aquatique est connu pour son intelligence et est dresse dans des parcs marins ?", 'da': 'Hvilket stort vandlevende pattedyr er kendt for sin intelligens og trænes i marineparker?'},
 'opt1': {'en': 'Dolphin', 'de': 'Delfin', 'es': 'Delfin', 'nl': 'Dolfijn', 'sv': 'Delfin', 'fr': 'Dauphin', 'da': 'Delfin'},
 'opt2': {'en': 'Orca', 'de': 'Orca', 'es': 'Orca', 'nl': 'Orka', 'sv': 'Sparval', 'fr': 'Orque', 'da': 'Spækhugger'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which country is home to the ancient city of Petra's Treasury facade?", 'de': 'In welchem Land liegt die beruehmte Schatzhausfassade von Petra?', 'es': '¿En que pais esta la famosa fachada del Tesoro de Petra?', 'nl': 'In welk land ligt de beroemde Treasury-gevel van Petra?', 'sv': 'I vilket land ligger Petras beromda Treasury-fasad?', 'fr': "Dans quel pays se trouve la celebre facade du Tresor de Petra ?", 'da': 'Hvilket land huser den berømte Treasury-facade i Petra?'},
 'opt1': {'en': 'Israel', 'de': 'Israel', 'es': 'Israel', 'nl': 'Israel', 'sv': 'Israel', 'fr': 'Israel', 'da': 'Israel'},
 'opt2': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which type of shoe, worn by ballet dancers, has a hard block toe for dancing en pointe?', 'de': 'Welche Art von Schuh, von Ballerinas getragen, hat eine harte Blockspitze zum Tanzen auf Spitze?', 'es': '¿Que tipo de zapato, usado por bailarinas de ballet, tiene punta dura para bailar en puntas?', 'nl': 'Welk type schoen, gedragen door balletdansers, heeft een harde blokneus om op spitzen te dansen?', 'sv': 'Vilken typ av sko, som baers av balettdansare, har en hard blocktott for att dansa pa toa?', 'fr': "Quel type de chaussure, portee par les danseuses de ballet, a un bout renforce pour danser en pointe ?", 'da': 'Hvilken type sko, båret af balletdansere, har en hård blokspids til at danse på spidsko?'},
 'opt1': {'en': 'Ballet flat', 'de': 'Ballerina (Schuh)', 'es': 'Bailarina (calzado)', 'nl': 'Balletschoen', 'sv': 'Ballerinasko', 'fr': 'Ballerine (chaussure)', 'da': 'Ballerinasko'},
 'opt2': {'en': 'Pointe shoe', 'de': 'Spitzenschuh', 'es': 'Zapatilla de punta', 'nl': 'Spitsschoen', 'sv': 'Toskor', 'fr': 'Chausson de pointe', 'da': 'Spidsko'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which streaming series follows a chess prodigy's rise through 1960s tournaments?", 'de': "Welche Streaming-Serie folgt dem Aufstieg eines Schachwunderkinds durch Turniere der 1960er Jahre?", 'es': "¿Que serie de streaming sigue el ascenso de un prodigio del ajedrez a traves de torneos de los años 60?", 'nl': "Welke streamingserie volgt de opkomst van een schaakwonderkind door toernooien in de jaren 60?", 'sv': "Vilken streamingserie foljer ett schackgenis uppgang genom 1960-talets turneringar?", 'fr': "Quelle serie en streaming suit l'ascension d'un prodige des echecs a travers des tournois des annees 1960 ?", 'da': 'Hvilken streamingserie følger et skakgenis vej gennem 1960’ernes turneringer?'},
 'opt1': {'en': 'Beautiful Mind', 'de': 'A Beautiful Mind', 'es': 'Una mente maravillosa', 'nl': 'A Beautiful Mind', 'sv': 'A Beautiful Mind', 'fr': 'Un homme d\'exception', 'da': 'A Beautiful Mind'},
 'opt2': {'en': "The Queen's Gambit", 'de': 'Das Damengambit', 'es': 'Gambito de dama', 'nl': "The Queen's Gambit", 'sv': 'Damgambit', 'fr': 'Le Jeu de la Dame', 'da': 'Dronninggambit'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which flower is known for its layered petals and is a classic symbol of romance in red?', 'de': 'Welche Blume ist fuer ihre geschichteten Bluetenblaetter bekannt und ist in Rot ein klassisches Symbol der Romantik?', 'es': '¿Que flor es conocida por sus petalos en capas y es un simbolo clasico del romance en rojo?', 'nl': 'Welke bloem staat bekend om zijn gelaagde bloemblaadjes en is in het rood een klassiek symbool van romantiek?', 'sv': 'Vilken blomma ar kand for sina lagerformade kronblad och ar en klassisk romantiksymbol i rott?', 'fr': "Quelle fleur est connue pour ses petales superposes et est un symbole classique de romance en rouge ?", 'da': 'Hvilken blomst er kendt for sine lagdelte kronblade og er et klassisk symbol på romantik i rødt?'},
 'opt1': {'en': 'Rose', 'de': 'Rose', 'es': 'Rosa', 'nl': 'Roos', 'sv': 'Ros', 'fr': 'Rose', 'da': 'Rose'},
 'opt2': {'en': 'Carnation', 'de': 'Nelke', 'es': 'Clavel', 'nl': 'Anjer', 'sv': 'Nejlika', 'fr': 'Oeillet', 'da': 'Nellike'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a rich, creamy potato salad often served at barbecues?', 'de': 'Welches suddstaatliche Gericht ist ein reichhaltiger, cremiger Kartoffelsalat, oft bei Grillfesten serviert?', 'es': '¿Que plato sureño de EE.UU. es una ensalada de papas rica y cremosa a menudo servida en barbacoas?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een rijke, romige aardappelsalade vaak geserveerd bij barbecues?', 'sv': 'Vilken sydstatlig ratt ar en rik, cremig potatissallad ofta serverad pa grillfester?', 'fr': "Quel plat du Sud des États-Unis est une salade de pommes de terre riche et cremeuse souvent servie lors de barbecues ?", 'da': 'Hvilken sydstatsret er en rig, cremet kartoffelsalat ofte serveret ved grillfester?'},
 'opt1': {'en': 'Potato salad', 'de': 'Kartoffelsalat', 'es': 'Ensalada de papas', 'nl': 'Aardappelsalade', 'sv': 'Potatissallad', 'fr': 'La salade de pommes de terre', 'da': 'Kartoffelsalat'},
 'opt2': {'en': 'Coleslaw', 'de': 'Krautsalat', 'es': 'Ensalada de col', 'nl': 'Koolsalade', 'sv': 'Coleslaw', 'fr': 'Salade de chou', 'da': 'Coleslaw'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which vitamin is essential for blood clotting and found in leafy greens like spinach?', 'de': 'Welches Vitamin ist essentiell fuer die Blutgerinnung und in Blattgemuese wie Spinat enthalten?', 'es': '¿Que vitamina es esencial para la coagulacion de la sangre y se encuentra en verduras de hoja como la espinaca?', 'nl': 'Welk vitamine is essentieel voor bloedstolling en te vinden in bladgroenten zoals spinazie?', 'sv': 'Vilket vitamin ar viktigt for blodkoagulering och finns i bladgronsaker som spenat?', 'fr': "Quelle vitamine est essentielle a la coagulation sanguine et se trouve dans les legumes-feuilles comme les epinards ?", 'da': 'Hvilket vitamin er essentielt for blodkoagulation og findes i bladgrønt som spinat?'},
 'opt1': {'en': 'Vitamin K', 'de': 'Vitamin K', 'es': 'Vitamina K', 'nl': 'Vitamine K', 'sv': 'Vitamin K', 'fr': 'Vitamine K', 'da': 'Vitamin K'},
 'opt2': {'en': 'Vitamin C', 'de': 'Vitamin C', 'es': 'Vitamina C', 'nl': 'Vitamine C', 'sv': 'Vitamin C', 'fr': 'Vitamine C', 'da': 'Vitamin C'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Volubilis?', 'de': 'In welchem Land liegt die antike Stadt Volubilis?', 'es': '¿En que pais esta la antigua ciudad de Volubilis?', 'nl': 'In welk land ligt de oude stad Volubilis?', 'sv': 'I vilket land ligger den antika staden Volubilis?', 'fr': "Dans quel pays se trouve l'ancienne cite de Volubilis ?", 'da': 'Hvilket land huser den antikke by Volubilis?'},
 'opt1': {'en': 'Algeria', 'de': 'Algerien', 'es': 'Argelia', 'nl': 'Algerije', 'sv': 'Algeriet', 'fr': "l'Algerie", 'da': 'Algeriet'},
 'opt2': {'en': 'Morocco', 'de': 'Marokko', 'es': 'Marruecos', 'nl': 'Marokko', 'sv': 'Marocko', 'fr': 'le Maroc', 'da': 'Marokko'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was known for hosting the first White House Easter Egg Roll?', 'de': 'Welche First Lady war bekannt fuer die Ausrichtung des ersten Osterei-Rollens im Weissen Haus?', 'es': '¿Que primera dama fue conocida por organizar la primera busqueda de huevos de Pascua en la Casa Blanca?', 'nl': 'Welke First Lady stond bekend om het organiseren van de eerste Easter Egg Roll in het Witte Huis?', 'sv': 'Vilken First Lady var kand for att vara vard for den forsta paskaggsrullningen i Vita huset?', 'fr': "Quelle Première dame etait connue pour avoir organise la premiere course aux oeufs de Paques a la Maison-Blanche ?", 'da': 'Hvilken førstedame var kendt for at være vært for det første Easter Egg Roll i Det Hvide Hus?'},
 'opt1': {'en': 'Lucy Hayes', 'de': 'Lucy Hayes', 'es': 'Lucy Hayes', 'nl': 'Lucy Hayes', 'sv': 'Lucy Hayes', 'fr': 'Lucy Hayes', 'da': 'Lucy Hayes'},
 'opt2': {'en': 'Julia Grant', 'de': 'Julia Grant', 'es': 'Julia Grant', 'nl': 'Julia Grant', 'sv': 'Julia Grant', 'fr': 'Julia Grant', 'da': 'Julia Grant'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which composer wrote 'Symphony No. 9' featuring the 'Ode to Joy'?", 'de': "Welcher Komponist schrieb die 'Sinfonie Nr. 9' mit der 'Ode an die Freude'?", 'es': "¿Que compositor escribio la 'Sinfonia No. 9' con la 'Oda a la alegria'?", 'nl': "Welke componist schreef 'Symfonie nr. 9' met de 'Ode aan de vreugde'?", 'sv': "Vilken kompositor skrev 'Symfoni nr 9' med 'Oden till gladjen'?", 'fr': "Quel compositeur a ecrit la 'Symphonie no 9' avec l'Ode a la joie ?", 'da': "Hvilken komponist skrev 'Symfoni nr. 9' med 'Ode til glæden'?"},
 'opt1': {'en': 'Ludwig van Beethoven', 'de': 'Ludwig van Beethoven', 'es': 'Ludwig van Beethoven', 'nl': 'Ludwig van Beethoven', 'sv': 'Ludwig van Beethoven', 'fr': 'Ludwig van Beethoven', 'da': 'Ludwig van Beethoven'},
 'opt2': {'en': 'Johannes Brahms', 'de': 'Johannes Brahms', 'es': 'Johannes Brahms', 'nl': 'Johannes Brahms', 'sv': 'Johannes Brahms', 'fr': 'Johannes Brahms', 'da': 'Johannes Brahms'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Elle Woods' friend Paulette in 'Legally Blonde'?", 'de': "Welche Schauspielerin spielte Elle Woods' Freundin Paulette in 'Natuerlich blond'?", 'es': "¿Que actriz interpreto a Paulette, amiga de Elle Woods, en 'Legalmente rubia'?", 'nl': "Welke actrice speelde Elle Woods' vriendin Paulette in 'Legally Blonde'?", 'sv': "Vilken skadespelerska spelade Elle Woods van Paulette i 'Blond makt'?", 'fr': "Quelle actrice a joue Paulette, l'amie d'Elle Woods, dans 'La Revanche d'une blonde' ?", 'da': "Hvilken skuespiller spillede Elle Woods' ven Paulette i 'Naturligt blond'?"},
 'opt1': {'en': 'Selma Blair', 'de': 'Selma Blair', 'es': 'Selma Blair', 'nl': 'Selma Blair', 'sv': 'Selma Blair', 'fr': 'Selma Blair', 'da': 'Selma Blair'},
 'opt2': {'en': 'Jennifer Coolidge', 'de': 'Jennifer Coolidge', 'es': 'Jennifer Coolidge', 'nl': 'Jennifer Coolidge', 'sv': 'Jennifer Coolidge', 'fr': 'Jennifer Coolidge', 'da': 'Jennifer Coolidge'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Southern U.S. dish is a slow-smoked meat, often pork, served with barbecue sauce?', 'de': 'Welches suddstaatliche Gericht ist langsam geraeuchertes Fleisch, oft Schwein, serviert mit Barbecue-Sauce?', 'es': '¿Que plato sureño de EE.UU. es carne ahumada lentamente, a menudo cerdo, servida con salsa barbacoa?', 'nl': 'Welk zuidelijk Amerikaans gerecht is langzaam gerookt vlees, vaak varkensvlees, geserveerd met barbecuesaus?', 'sv': 'Vilken sydstatlig ratt ar langsamt rokt kott, ofta flask, serverad med barbecuesas?', 'fr': "Quel plat du Sud des États-Unis est une viande fumee lentement, souvent du porc, servie avec de la sauce barbecue ?", 'da': 'Hvilken sydstatsret er langsomt røget kød, ofte svinekød, serveret med barbecuesovs?'},
 'opt1': {'en': 'Brisket', 'de': 'Brisket', 'es': 'Brisket', 'nl': 'Brisket', 'sv': 'Brisket', 'fr': 'Poitrine de boeuf', 'da': 'Brisket'},
 'opt2': {'en': 'Pulled pork', 'de': 'Pulled Pork', 'es': 'Cerdo desmenuzado', 'nl': 'Pulled pork', 'sv': 'Pulled pork', 'fr': 'Porc effiloche', 'da': 'Pulled pork'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which U.S. state is home to the Hoover Dam?', 'de': 'In welchem US-Bundesstaat liegt der Hoover-Staudamm?', 'es': '¿En que estado de EE.UU. esta la Presa Hoover?', 'nl': 'In welke Amerikaanse staat ligt de Hoover Dam?', 'sv': 'I vilken amerikansk delstat ligger Hooverdammen?', 'fr': "Dans quel État americain se trouve le barrage Hoover ?", 'da': 'I hvilken amerikansk delstat ligger Hoover-dæmningen?'},
 'opt1': {'en': 'Arizona', 'de': 'Arizona', 'es': 'Arizona', 'nl': 'Arizona', 'sv': 'Arizona', 'fr': "l'Arizona", 'da': 'Arizona'},
 'opt2': {'en': 'Nevada', 'de': 'Nevada', 'es': 'Nevada', 'nl': 'Nevada', 'sv': 'Nevada', 'fr': 'le Nevada', 'da': 'Nevada'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which small mammal has webbed feet and a flat tail, known for slapping water as a warning?', 'de': 'Welches kleine Saeugetier hat Schwimmhaeute und einen flachen Schwanz, bekannt fuer das Wasserschlagen als Warnung?', 'es': '¿Que pequeño mamifero tiene patas palmeadas y una cola plana, conocido por golpear el agua como advertencia?', 'nl': 'Welk klein zoogdier heeft zwemvliezen en een platte staart, bekend om het slaan van water als waarschuwing?', 'sv': 'Vilket litet daggdjur har simhud och en platt svans, kant for att slå vattnet som varning?', 'fr': "Quel petit mammifere a des pieds palmes et une queue plate, connu pour claquer l'eau en guise d'avertissement ?", 'da': 'Hvilket lille pattedyr har svømmefødder og en flad hale, kendt for at slå vandet som en advarsel?'},
 'opt1': {'en': 'Beaver', 'de': 'Biber', 'es': 'Castor', 'nl': 'Bever', 'sv': 'Bavern', 'fr': 'Castor', 'da': 'Bæver'},
 'opt2': {'en': 'Otter', 'de': 'Otter', 'es': 'Nutria', 'nl': 'Otter', 'sv': 'Utter', 'fr': 'Loutre', 'da': 'Odder'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Jewish tradition involves a ceremonial meal to mark the beginning of Passover?', 'de': 'Welche juedische Tradition beinhaltet eine zeremonielle Mahlzeit zum Beginn des Passahfests?', 'es': '¿Que tradicion judia implica una comida ceremonial para marcar el inicio de la Pascua?', 'nl': 'Welke Joodse traditie omvat een ceremoniele maaltijd om het begin van Pesach te markeren?', 'sv': 'Vilken judisk tradition innebar en ceremoniell maltid for att markera borjan av pask (Pesach)?', 'fr': "Quelle tradition juive implique un repas ceremoniel pour marquer le debut de Pessa'h ?", 'da': 'Hvilken jødisk tradition involverer et ceremonielt måltid for at markere begyndelsen af påske (pesach)?'},
 'opt1': {'en': 'Seder', 'de': 'Seder', 'es': 'Seder', 'nl': 'Seideravond', 'sv': 'Seder', 'fr': 'Le seder', 'da': 'Seder'},
 'opt2': {'en': 'Shabbat dinner', 'de': 'Schabbat-Abendessen', 'es': 'Cena de Shabat', 'nl': 'Sjabbatmaaltijd', 'sv': 'Shabbatmiddag', 'fr': 'Le diner de Shabbat', 'da': 'Sabbatsmiddag'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'Pride and Prejudice'?", 'de': "Welche Autorin schrieb 'Stolz und Vorurteil'?", 'es': "¿Que autora escribio 'Orgullo y prejuicio'?", 'nl': "Welke auteur schreef 'Trots en Vooroordeel'?", 'sv': "Vilken forfattare skrev 'Stolthet och fordom'?", 'fr': "Quelle autrice a ecrit 'Orgueil et Prejuges' ?", 'da': "Hvilken forfatter skrev 'Stolthed og fordom'?"},
 'opt1': {'en': 'George Eliot', 'de': 'George Eliot', 'es': 'George Eliot', 'nl': 'George Eliot', 'sv': 'George Eliot', 'fr': 'George Eliot', 'da': 'George Eliot'},
 'opt2': {'en': 'Jane Austen', 'de': 'Jane Austen', 'es': 'Jane Austen', 'nl': 'Jane Austen', 'sv': 'Jane Austen', 'fr': 'Jane Austen', 'da': 'Jane Austen'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Petra, a UNESCO World Heritage Site?', 'de': 'In welchem Land liegt die antike Stadt Petra, ein UNESCO-Weltkulturerbe?', 'es': '¿En que pais esta la antigua ciudad de Petra, Patrimonio de la Humanidad de la UNESCO?', 'nl': 'In welk land ligt de oude stad Petra, een UNESCO-werelderfgoedgebied?', 'sv': 'I vilket land ligger den antika staden Petra, ett varldsarv?', 'fr': "Dans quel pays se trouve l'ancienne cite de Petra, site du patrimoine mondial de l'UNESCO ?", 'da': 'Hvilket land huser den antikke by Petra, et UNESCO-verdensarvssted?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Egypt', 'de': 'Ägypten', 'es': 'Egipto', 'nl': 'Egypte', 'sv': 'Egypten', 'fr': "l'Egypte", 'da': 'Egypten'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which instrument is central to traditional Chinese music, a plucked string instrument?', 'de': 'Welches Instrument ist zentral fuer traditionelle chinesische Musik, ein gezupftes Saiteninstrument?', 'es': '¿Que instrumento es central en la musica china tradicional, un instrumento de cuerda punteada?', 'nl': 'Welk instrument is centraal in traditionele Chinese muziek, een getokkeld snaarinstrument?', 'sv': 'Vilket instrument ar centralt i traditionell kinesisk musik, ett knappt strakinstrument?', 'fr': "Quel instrument est central dans la musique chinoise traditionnelle, un instrument a cordes pincees ?", 'da': 'Hvilket instrument er centralt i traditionel kinesisk musik, et klimpret strengeinstrument?'},
 'opt1': {'en': 'Guzheng', 'de': 'Guzheng', 'es': 'Guzheng', 'nl': 'Guzheng', 'sv': 'Guzheng', 'fr': 'Guzheng', 'da': 'Guzheng'},
 'opt2': {'en': 'Erhu', 'de': 'Erhu', 'es': 'Erhu', 'nl': 'Erhu', 'sv': 'Erhu', 'fr': 'Erhu', 'da': 'Erhu'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a layered dessert with pudding, bananas, and vanilla wafers?', 'de': 'Welches suddstaatliche Gericht ist ein Schichtdessert mit Pudding, Bananen und Vanillewaffeln?', 'es': '¿Que plato sureño de EE.UU. es un postre en capas con pudin, platanos y galletas de vainilla?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een gelaagd dessert met pudding, bananen en vanillewafeltjes?', 'sv': 'Vilken sydstatlig ratt ar en lagerdessert med pudding, bananer och vaniljkex?', 'fr': "Quel plat du Sud des États-Unis est un dessert en couches avec du pudding, des bananes et des gaufrettes a la vanille ?", 'da': 'Hvilken sydstatsret er en lagdelt dessert med budding, bananer og vaniljekiks?'},
 'opt1': {'en': 'Banana pudding', 'de': 'Bananenpudding', 'es': 'Pudin de platano', 'nl': 'Bananenpudding', 'sv': 'Bananpudding', 'fr': 'Pudding a la banane', 'da': 'Bananbudding'},
 'opt2': {'en': 'Trifle', 'de': 'Trifle', 'es': 'Trifle', 'nl': 'Trifle', 'sv': 'Trifle', 'fr': 'Trifle', 'da': 'Trifli'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which mineral is essential for oxygen transport in red blood cells?', 'de': 'Welches Mineral ist essentiell fuer den Sauerstofftransport in roten Blutkoerperchen?', 'es': '¿Que mineral es esencial para el transporte de oxigeno en los globulos rojos?', 'nl': 'Welk mineraal is essentieel voor zuurstoftransport in rode bloedcellen?', 'sv': 'Vilket mineral ar viktigt for syretransport i roda blodkroppar?', 'fr': "Quel mineral est essentiel au transport de l'oxygene dans les globules rouges ?", 'da': 'Hvilket mineral er essentielt for iltransport i røde blodlegemer?'},
 'opt1': {'en': 'Iron', 'de': 'Eisen', 'es': 'Hierro', 'nl': 'IJzer', 'sv': 'Jarn', 'fr': 'Fer', 'da': 'Jern'},
 'opt2': {'en': 'Zinc', 'de': 'Zink', 'es': 'Zinc', 'nl': 'Zink', 'sv': 'Zink', 'fr': 'Zinc', 'da': 'Zink'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which film has Julia Roberts playing a woman who escapes an abusive marriage?', 'de': 'In welchem Film spielt Julia Roberts eine Frau, die aus einer missbrauchlichen Ehe flieht?', 'es': '¿En que pelicula Julia Roberts interpreta a una mujer que escapa de un matrimonio abusivo?', 'nl': 'In welke film speelt Julia Roberts een vrouw die ontsnapt aan een gewelddadig huwelijk?', 'sv': 'I vilken film spelar Julia Roberts en kvinna som flyr fran ett vald samt aktenskap?', 'fr': "Dans quel film Julia Roberts joue-t-elle une femme qui echappe a un mariage abusif ?", 'da': 'I hvilken film spiller Julia Roberts en kvinde, der flygter fra et voldeligt ægteskab?'},
 'opt1': {'en': 'The Pelican Brief', 'de': 'Die Akte', 'es': 'El informe pelicano', 'nl': 'The Pelican Brief', 'sv': 'Pelikanbrevet', 'fr': "L'Affaire Pelican", 'da': 'The Pelican Brief'},
 'opt2': {'en': 'Sleeping with the Enemy', 'de': 'Schlafen mit dem Feind', 'es': 'Durmiendo con su enemigo', 'nl': 'Sleeping with the Enemy', 'sv': 'Sova med fienden', 'fr': "Dors avec tes yeux", 'da': 'Sleeping with the Enemy'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which country is home to the ancient city of Petra's neighbor, the Red Sea resort of Aqaba?", 'de': 'In welchem Land liegt der Rotmeer-Badeort Aqaba, nahe Petra?', 'es': '¿En que pais esta el balneario de Aqaba en el Mar Rojo, cerca de Petra?', 'nl': 'In welk land ligt de Rode Zee-badplaats Aqaba, nabij Petra?', 'sv': 'I vilket land ligger Rodahavsorten Aqaba, nara Petra?', 'fr': "Dans quel pays se trouve la station balneaire d'Aqaba sur la mer Rouge, pres de Petra ?", 'da': 'Hvilket land huser Det Røde Hav-resortet Aqaba, nær Petra?'},
 'opt1': {'en': 'Egypt', 'de': 'Ägypten', 'es': 'Egipto', 'nl': 'Egypte', 'sv': 'Egypten', 'fr': "l'Egypte", 'da': 'Egypten'},
 'opt2': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of jewelry, worn around the neck, often has a pendant hanging from a chain?', 'de': 'Welche Art von Schmuck, um den Hals getragen, hat oft einen Anhaenger an einer Kette?', 'es': '¿Que tipo de joya, llevada alrededor del cuello, a menudo tiene un colgante que cuelga de una cadena?', 'nl': 'Welk type sieraad, om de nek gedragen, heeft vaak een hanger aan een ketting?', 'sv': 'Vilken typ av smycke, som baers runt halsen, har ofta ett hangsmycke som hanger i en kedja?', 'fr': "Quel type de bijou, porte autour du cou, a souvent un pendentif accroche a une chaine ?", 'da': 'Hvilken type smykke, båret om halsen, har ofte en vedhæng, der hænger fra en kæde?'},
 'opt1': {'en': 'Necklace', 'de': 'Halskette', 'es': 'Collar', 'nl': 'Ketting', 'sv': 'Halsband', 'fr': 'Collier', 'da': 'Halskæde'},
 'opt2': {'en': 'Choker', 'de': 'Choker', 'es': 'Gargantilla', 'nl': 'Choker', 'sv': 'Choker', 'fr': 'Ras-de-cou', 'da': 'Choker'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which small, brightly colored bird is known for hovering in place and drinking nectar from flowers?', 'de': 'Welcher kleine, bunte Vogel ist dafuer bekannt, auf der Stelle zu schweben und Nektar aus Blueten zu trinken?', 'es': '¿Que pequeña ave de colores brillantes es conocida por planear en el aire y beber nectar de las flores?', 'nl': 'Welke kleine, felgekleurde vogel staat bekend om het stilstaand zweven en het drinken van nectar uit bloemen?', 'sv': 'Vilken liten, klart fargad fagel ar kand for att sväva pa stallet och dricka nektar fran blommor?', 'fr': "Quel petit oiseau vivement colore est connu pour faire du surplace et boire le nectar des fleurs ?", 'da': 'Hvilken lille, klart farvet fugl er kendt for at svæve på stedet og drikke nektar fra blomster?'},
 'opt1': {'en': 'Sunbird', 'de': 'Nektarvogel', 'es': 'Ave del sol', 'nl': 'Zonnevogel', 'sv': 'Solfagel', 'fr': 'Souimanga', 'da': 'Solfugl'},
 'opt2': {'en': 'Hummingbird', 'de': 'Kolibri', 'es': 'Colibri', 'nl': 'Kolibrie', 'sv': 'Kolibri', 'fr': 'Colibri', 'da': 'Kolibri'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which First Lady was the first to travel abroad extensively as a diplomatic representative?', 'de': 'Welche First Lady reiste als erste ausgiebig ins Ausland als diplomatische Vertreterin?', 'es': '¿Que primera dama fue la primera en viajar extensamente al extranjero como representante diplomatica?', 'nl': 'Welke First Lady reisde als eerste uitgebreid naar het buitenland als diplomatiek vertegenwoordiger?', 'sv': 'Vilken First Lady var forst med att resa mycket utomlands som diplomatisk representant?', 'fr': "Quelle Première dame a ete la premiere a voyager abondamment a l'etranger en tant que representante diplomatique ?", 'da': 'Hvilken førstedame var den første til at rejse omfattende i udlandet som diplomatisk repræsentant?'},
 'opt1': {'en': 'Pat Nixon', 'de': 'Pat Nixon', 'es': 'Pat Nixon', 'nl': 'Pat Nixon', 'sv': 'Pat Nixon', 'fr': 'Pat Nixon', 'da': 'Pat Nixon'},
 'opt2': {'en': 'Eleanor Roosevelt', 'de': 'Eleanor Roosevelt', 'es': 'Eleanor Roosevelt', 'nl': 'Eleanor Roosevelt', 'sv': 'Eleanor Roosevelt', 'fr': 'Eleanor Roosevelt', 'da': 'Eleanor Roosevelt'}},

{'category': 'History', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which First Lady was a former librarian who championed literacy programs nationwide?", 'de': 'Welche First Lady war eine ehemalige Bibliothekarin, die sich landesweit fuer Alphabetisierungsprogramme einsetzte?', 'es': '¿Que primera dama fue una ex bibliotecaria que defendio los programas de alfabetizacion a nivel nacional?', 'nl': 'Welke First Lady was een voormalige bibliothecaris die zich landelijk inzette voor geletterdheidsprogramma\'s?', 'sv': 'Vilken First Lady var en tidigare bibliotekarie som forespakade lasfrämjandeprogram over hela landet?', 'fr': "Quelle Première dame etait une ancienne bibliothecaire qui a defendu des programmes d'alphabetisation a l'echelle nationale ?", 'da': 'Hvilken førstedame var en tidligere bibliotekar, der gik i bresche for læsefærdighedsprogrammer landsdækkende?'},
 'opt1': {'en': 'Barbara Bush', 'de': 'Barbara Bush', 'es': 'Barbara Bush', 'nl': 'Barbara Bush', 'sv': 'Barbara Bush', 'fr': 'Barbara Bush', 'da': 'Barbara Bush'},
 'opt2': {'en': 'Laura Bush', 'de': 'Laura Bush', 'es': 'Laura Bush', 'nl': 'Laura Bush', 'sv': 'Laura Bush', 'fr': 'Laura Bush', 'da': 'Laura Bush'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which Egyptian author won the Nobel Prize in Literature in 1988?', 'de': 'Welcher aegyptische Autor gewann 1988 den Nobelpreis fuer Literatur?', 'es': '¿Que autor egipcio gano el Premio Nobel de Literatura en 1988?', 'nl': 'Welke Egyptische auteur won in 1988 de Nobelprijs voor Literatuur?', 'sv': 'Vilken egyptisk forfattare vann Nobelpriset i litteratur 1988?', 'fr': "Quel auteur egyptien a remporte le prix Nobel de litterature en 1988 ?", 'da': 'Hvilken egyptisk forfatter vandt Nobelprisen i litteratur i 1988?'},
 'opt1': {'en': 'Taha Hussein', 'de': 'Taha Hussein', 'es': 'Taha Husein', 'nl': 'Taha Hoessein', 'sv': 'Taha Hussein', 'fr': 'Taha Hussein', 'da': 'Taha Hussein'},
 'opt2': {'en': 'Naguib Mahfouz', 'de': 'Nagib Mahfuz', 'es': 'Naguib Mahfuz', 'nl': 'Naguib Mahfoez', 'sv': 'Naguib Mahfouz', 'fr': 'Naguib Mahfouz', 'da': 'Naguib Mahfouz'}},

{'category': 'Science', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which vitamin, produced by sun exposure, helps the body absorb calcium?', 'de': 'Welches Vitamin, durch Sonnenexposition produziert, hilft dem Koerper, Kalzium aufzunehmen?', 'es': '¿Que vitamina, producida por la exposicion al sol, ayuda al cuerpo a absorber calcio?', 'nl': 'Welk vitamine, geproduceerd door blootstelling aan zon, helpt het lichaam calcium op te nemen?', 'sv': 'Vilket vitamin, som produceras av solexponering, hjalper kroppen att absorbera kalcium?', 'fr': "Quelle vitamine, produite par l'exposition au soleil, aide le corps a absorber le calcium ?", 'da': 'Hvilket vitamin, produceret ved soleksponering, hjælper kroppen med at optage calcium?'},
 'opt1': {'en': 'Vitamin B6', 'de': 'Vitamin B6', 'es': 'Vitamina B6', 'nl': 'Vitamine B6', 'sv': 'Vitamin B6', 'fr': 'Vitamine B6', 'da': 'Vitamin B6'},
 'opt2': {'en': 'Vitamin D', 'de': 'Vitamin D', 'es': 'Vitamina D', 'nl': 'Vitamine D', 'sv': 'Vitamin D', 'fr': 'Vitamine D', 'da': 'Vitamin D'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Norwegian expressionist painter is known for 'The Scream'?", 'de': "Welcher norwegische expressionistische Maler ist fuer 'Der Schrei' bekannt?", 'es': "¿Que pintor expresionista noruego es conocido por 'El grito'?", 'nl': "Welke Noorse expressionistische schilder staat bekend om 'De Schreeuw'?", 'sv': "Vilken norsk expressionistisk malare ar kand for 'Skriet'?", 'fr': "Quel peintre expressionniste norvegien est connu pour 'Le Cri' ?", 'da': "Hvilken norsk ekspressionistisk maler er kendt for 'Skriget'?"},
 'opt1': {'en': 'Edvard Munch', 'de': 'Edvard Munch', 'es': 'Edvard Munch', 'nl': 'Edvard Munch', 'sv': 'Edvard Munch', 'fr': 'Edvard Munch', 'da': 'Edvard Munch'},
 'opt2': {'en': 'Akseli Gallen-Kallela', 'de': 'Akseli Gallen-Kallela', 'es': 'Akseli Gallen-Kallela', 'nl': 'Akseli Gallen-Kallela', 'sv': 'Akseli Gallen-Kallela', 'fr': 'Akseli Gallen-Kallela', 'da': 'Akseli Gallen-Kallela'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Southeast Asian country's capital is Vientiane?", 'de': 'Welches suedostasiatische Land hat Vientiane als Hauptstadt?', 'es': '¿Que pais del sudeste asiatico tiene como capital a Vientian?', 'nl': 'Welk Zuidoost-Aziatisch land heeft Vientiane als hoofdstad?', 'sv': 'Vilket sydostasiatiskt lands huvudstad ar Vientiane?', 'fr': "Quel pays d'Asie du Sud-Est a pour capitale Vientiane ?", 'da': 'Hvilket sydøstasiatisk lands hovedstad er Vientiane?'},
 'opt1': {'en': 'Laos', 'de': 'Laos', 'es': 'Laos', 'nl': 'Laos', 'sv': 'Laos', 'fr': 'le Laos', 'da': 'Laos'},
 'opt2': {'en': 'Cambodia', 'de': 'Kambodscha', 'es': 'Camboya', 'nl': 'Cambodja', 'sv': 'Kambodja', 'fr': 'le Cambodge', 'da': 'Cambodja'}},

{'category': 'Music', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Polish composer is known for his nocturnes and mazurkas for piano?', 'de': 'Welcher polnische Komponist ist fuer seine Nocturnes und Mazurkas fuer Klavier bekannt?', 'es': '¿Que compositor polaco es conocido por sus nocturnos y mazurcas para piano?', 'nl': 'Welke Poolse componist staat bekend om zijn nocturnes en mazurka\'s voor piano?', 'sv': 'Vilken polsk kompositor ar kand for sina nocturner och mazurkor for piano?', 'fr': "Quel compositeur polonais est connu pour ses nocturnes et mazurkas pour piano ?", 'da': 'Hvilken polsk komponist er kendt for sine nocturner og mazurkaer for klaver?'},
 'opt1': {'en': 'Frederic Chopin', 'de': 'Frederic Chopin', 'es': 'Frederic Chopin', 'nl': 'Frederic Chopin', 'sv': 'Frederic Chopin', 'fr': 'Frederic Chopin', 'da': 'Frederic Chopin'},
 'opt2': {'en': 'Franz Liszt', 'de': 'Franz Liszt', 'es': 'Franz Liszt', 'nl': 'Franz Liszt', 'sv': 'Franz Liszt', 'fr': 'Franz Liszt', 'da': 'Franz Liszt'}},

{'category': 'Food', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which Italian dish is a rice dish slowly cooked with broth, becoming creamy?', 'de': 'Welches italienische Gericht ist ein Reisgericht, langsam mit Bruehe gekocht, das cremig wird?', 'es': '¿Que plato italiano es un plato de arroz cocido lentamente con caldo, volviendose cremoso?', 'nl': 'Welk Italiaans gerecht is een rijstgerecht langzaam gekookt met bouillon, dat romig wordt?', 'sv': 'Vilken italiensk ratt ar en risratt langsamt tillagad med buljong, som blir cremig?', 'fr': "Quel plat italien est un plat de riz cuit lentement avec du bouillon, devenant cremeux ?", 'da': 'Hvilken italiensk ret er en risret langsomt tilberedt med bouillon, der bliver cremet?'},
 'opt1': {'en': 'Polenta', 'de': 'Polenta', 'es': 'Polenta', 'nl': 'Polenta', 'sv': 'Polenta', 'fr': 'Polenta', 'da': 'Polenta'},
 'opt2': {'en': 'Risotto', 'de': 'Risotto', 'es': 'Risotto', 'nl': 'Risotto', 'sv': 'Risotto', 'fr': 'Risotto', 'da': 'Risotto'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'In Greek mythology, which figure was doomed to roll a boulder up a hill for eternity?', 'de': 'In der griechischen Mythologie, welche Figur war dazu verdammt, ewig einen Felsbrocken einen Huegel hinaufzurollen?', 'es': '¿En la mitologia griega, que figura estaba condenada a rodar una roca cuesta arriba por la eternidad?', 'nl': 'In de Griekse mythologie, welke figuur was gedoemd om voor eeuwig een rotsblok een heuvel op te rollen?', 'sv': 'I grekisk mytologi, vilken figur var domd att rulla ett stenblock upp for en kulle for evigt?', 'fr': "Dans la mythologie grecque, quel personnage etait condamne a rouler un rocher en haut d'une colline pour l'eternite ?", 'da': 'I græsk mytologi, hvilken figur var dømt til at rulle en klippeblok op ad en bakke i al evighed?'},
 'opt1': {'en': 'Tantalus', 'de': 'Tantalos', 'es': 'Tantalo', 'nl': 'Tantalus', 'sv': 'Tantalos', 'fr': 'Tantale', 'da': 'Tantalos'},
 'opt2': {'en': 'Sisyphus', 'de': 'Sisyphos', 'es': 'Sisifo', 'nl': 'Sisyfus', 'sv': 'Sisyfos', 'fr': 'Sisyphe', 'da': 'Sisyfos'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which actress played the title blues singer in 'Ma Rainey's Black Bottom'?", 'de': "Welche Schauspielerin spielte die Blues-Saengerin in 'Ma Raineys schwarzer Hintern'?", 'es': "¿Que actriz interpreto a la cantante de blues protagonista en 'Ma Rainey's Black Bottom'?", 'nl': "Welke actrice speelde de titelblueszangeres in 'Ma Rainey's Black Bottom'?", 'sv': "Vilken skadespelerska spelade titelns bluessangerska i 'Ma Rainey's Black Bottom'?", 'fr': "Quelle actrice a joue la chanteuse de blues titulaire dans 'Ma Rainey's Black Bottom' ?", 'da': "Hvilken skuespiller spillede titelblues-sangeren i 'Ma Rainey's Black Bottom'?"},
 'opt1': {'en': 'Viola Davis', 'de': 'Viola Davis', 'es': 'Viola Davis', 'nl': 'Viola Davis', 'sv': 'Viola Davis', 'fr': 'Viola Davis', 'da': 'Viola Davis'},
 'opt2': {'en': 'Angela Bassett', 'de': 'Angela Bassett', 'es': 'Angela Bassett', 'nl': 'Angela Bassett', 'sv': 'Angela Bassett', 'fr': 'Angela Bassett', 'da': 'Angela Bassett'}},

{'category': 'Space', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which dwarf galaxy, visible to the naked eye from the Southern Hemisphere, is one of the Milky Way's closest neighbors?", 'de': 'Welche Zwerggalaxie, mit blossem Auge von der Suedhalbkugel aus sichtbar, ist einer der naechsten Nachbarn der Milchstrasse?', 'es': '¿Que galaxia enana, visible a simple vista desde el Hemisferio Sur, es una de las vecinas mas cercanas de la Via Lactea?', 'nl': 'Welk dwergstelsel, met het blote oog zichtbaar vanaf het zuidelijk halfrond, is een van de dichtstbijzijnde buren van de Melkweg?', 'sv': 'Vilken dvarggalax, synlig for blotta ogat fran sodra halvklotet, ar en av Vintergatans narmaste grannar?', 'fr': "Quelle galaxie naine, visible a l'oeil nu depuis l'hemisphere Sud, est l'un des plus proches voisins de la Voie lactee ?", 'da': 'Hvilken dværggalakse, synlig med det blotte øje fra den sydlige halvkugle, er en af Mælkevejens nærmeste naboer?'},
 'opt1': {'en': 'Large Magellanic Cloud', 'de': 'Große Magellansche Wolke', 'es': 'Gran Nube de Magallanes', 'nl': 'Grote Magelhaense Wolk', 'sv': 'Stora magellanska molnet', 'fr': 'Grand Nuage de Magellan', 'da': 'Store Magellanske Sky'},
 'opt2': {'en': 'Andromeda Galaxy', 'de': 'Andromedagalaxie', 'es': 'Galaxia de Andromeda', 'nl': 'Andromedastelsel', 'sv': 'Andromedagalaxen', 'fr': "Galaxie d'Andromede", 'da': 'Andromedagalaksen'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which African country is home to Table Mountain?', 'de': 'In welchem afrikanischen Land liegt der Tafelberg?', 'es': '¿En que pais africano esta la Montaña de la Mesa?', 'nl': 'In welk Afrikaans land ligt Tafelberg?', 'sv': 'I vilket afrikanskt land ligger Taffelberget?', 'fr': "Dans quel pays africain se trouve la montagne de la Table ?", 'da': 'Hvilket afrikansk land huser Table Mountain?'},
 'opt1': {'en': 'South Africa', 'de': 'Suedafrika', 'es': 'Sudafrica', 'nl': 'Zuid-Afrika', 'sv': 'Sydafrika', 'fr': "l'Afrique du Sud", 'da': 'Sydafrika'},
 'opt2': {'en': 'Namibia', 'de': 'Namibia', 'es': 'Namibia', 'nl': 'Namibie', 'sv': 'Namibia', 'fr': 'la Namibie', 'da': 'Namibia'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which American designer is known for minimalist, sportswear-inspired collections in the 1970s-90s?', 'de': 'Welcher amerikanische Designer ist bekannt fuer minimalistische, sportswear-inspirierte Kollektionen in den 1970er-90er Jahren?', 'es': '¿Que diseñador estadounidense es conocido por colecciones minimalistas inspiradas en ropa deportiva en los años 70-90?', 'nl': 'Welke Amerikaanse ontwerper staat bekend om minimalistische, door sportkleding geinspireerde collecties in de jaren 70-90?', 'sv': 'Vilken amerikansk designer ar kand for minimalistiska, sportklädesinspirerade kollektioner pa 1970-90-talen?', 'fr': "Quel createur americain est connu pour ses collections minimalistes inspirees du sportswear dans les annees 1970-90 ?", 'da': 'Hvilken amerikansk designer er kendt for minimalistiske, sportstøjsinspirerede kollektioner i 1970’erne-90’erne?'},
 'opt1': {'en': 'Ralph Lauren', 'de': 'Ralph Lauren', 'es': 'Ralph Lauren', 'nl': 'Ralph Lauren', 'sv': 'Ralph Lauren', 'fr': 'Ralph Lauren', 'da': 'Ralph Lauren'},
 'opt2': {'en': 'Calvin Klein', 'de': 'Calvin Klein', 'es': 'Calvin Klein', 'nl': 'Calvin Klein', 'sv': 'Calvin Klein', 'fr': 'Calvin Klein', 'da': 'Calvin Klein'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which Greek philosopher is known for the Socratic method of questioning?', 'de': 'Welcher griechische Philosoph ist fuer die sokratische Fragemethode bekannt?', 'es': '¿Que filosofo griego es conocido por el metodo socratico de preguntas?', 'nl': 'Welke Griekse filosoof staat bekend om de socratische vraagmethode?', 'sv': 'Vilken grekisk filosof ar kand for den sokratiska fragemetoden?', 'fr': "Quel philosophe grec est connu pour la methode socratique de questionnement ?", 'da': 'Hvilken græsk filosof er kendt for den sokratiske spørgemetode?'},
 'opt1': {'en': 'Plato', 'de': 'Platon', 'es': 'Platon', 'nl': 'Plato', 'sv': 'Platon', 'fr': 'Platon', 'da': 'Platon'},
 'opt2': {'en': 'Socrates', 'de': 'Sokrates', 'es': 'Socrates', 'nl': 'Socrates', 'sv': 'Sokrates', 'fr': 'Socrate', 'da': 'Sokrates'}},

{'category': 'History', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which quilting pattern uses a repeated basket shape pieced from triangles and squares?', 'de': 'Welches Quiltmuster verwendet eine wiederholte Korbform aus Dreiecken und Quadraten?', 'es': '¿Que patron de acolchado usa una forma repetida de canasta hecha de triangulos y cuadrados?', 'nl': 'Welk quiltpatroon gebruikt een herhaalde mandvorm gemaakt van driehoeken en vierkanten?', 'sv': 'Vilket lapptacksmonster anvander en upprepad korgform gjord av trianglar och fyrkanter?', 'fr': "Quel motif de patchwork utilise une forme repetee de panier assemblee a partir de triangles et de carres ?", 'da': 'Hvilket quiltmønster bruger en gentaget kurvformet figur sammensat af trekanter og firkanter?'},
 'opt1': {'en': 'Churn Dash', 'de': 'Churn Dash', 'es': 'Churn Dash', 'nl': 'Churn Dash', 'sv': 'Churn Dash', 'fr': 'Churn Dash', 'da': 'Churn Dash'},
 'opt2': {'en': 'Basket quilt', 'de': 'Korb-Quilt', 'es': 'Edredon de cesta', 'nl': 'Mandquilt', 'sv': 'Korgquilt', 'fr': 'Patchwork panier', 'da': 'Kurvequilt'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which flowering plant is known as 'queen of the night' for its rare, one-night bloom?", 'de': "Welche bluehende Pflanze wird 'Koenigin der Nacht' genannt, wegen ihrer seltenen, nur eine Nacht dauernden Bluete?", 'es': "¿Que planta con flores es conocida como 'reina de la noche' por su rara floracion de una sola noche?", 'nl': "Welke bloeiende plant staat bekend als 'koningin van de nacht' vanwege zijn zeldzame, eennachtsbloei?", 'sv': "Vilken blommande vaxt kallas 'nattens drottning' for sin sallsynta, en-natts-blomning?", 'fr': "Quelle plante a fleurs est connue comme la 'reine de la nuit' pour sa rare floraison d'une seule nuit ?", 'da': "Hvilken blomstrende plante er kendt som 'nattens dronning' for sin sjældne, en-nats blomstring?"},
 'opt1': {'en': 'Night-blooming cereus', 'de': 'Koenigin der Nacht (Kaktus)', 'es': 'Cactus reina de la noche', 'nl': 'Koningin van de nacht (cactus)', 'sv': 'Nattens drottning (kaktus)', 'fr': 'Cierge de nuit', 'da': 'Nattens dronning (kaktus)'},
 'opt2': {'en': 'Moonflower', 'de': 'Mondblume', 'es': 'Flor de luna', 'nl': 'Maanbloem', 'sv': 'Manblomma', 'fr': 'Fleur de lune', 'da': 'Måneblomst'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which architect designed the Bauhaus school building in Dessau, Germany?', 'de': 'Welcher Architekt entwarf das Bauhaus-Schulgebaeude in Dessau, Deutschland?', 'es': '¿Que arquitecto diseño el edificio de la escuela Bauhaus en Dessau, Alemania?', 'nl': 'Welke architect ontwierp het Bauhaus-schoolgebouw in Dessau, Duitsland?', 'sv': 'Vilken arkitekt designade Bauhausskolans byggnad i Dessau, Tyskland?', 'fr': "Quel architecte a conçu le batiment de l'ecole Bauhaus a Dessau, en Allemagne ?", 'da': 'Hvilken arkitekt designede Bauhaus-skolebygningen i Dessau, Tyskland?'},
 'opt1': {'en': 'Ludwig Mies van der Rohe', 'de': 'Ludwig Mies van der Rohe', 'es': 'Ludwig Mies van der Rohe', 'nl': 'Ludwig Mies van der Rohe', 'sv': 'Ludwig Mies van der Rohe', 'fr': 'Ludwig Mies van der Rohe', 'da': 'Ludwig Mies van der Rohe'},
 'opt2': {'en': 'Walter Gropius', 'de': 'Walter Gropius', 'es': 'Walter Gropius', 'nl': 'Walter Gropius', 'sv': 'Walter Gropius', 'fr': 'Walter Gropius', 'da': 'Walter Gropius'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Greek mythology, which god is the messenger of the gods, known for his winged sandals?', 'de': 'In der griechischen Mythologie, welcher Gott ist der Botenbringer der Goetter, bekannt fuer seine gefluegelten Sandalen?', 'es': '¿En la mitologia griega, que dios es el mensajero de los dioses, conocido por sus sandalias aladas?', 'nl': 'In de Griekse mythologie, welke god is de boodschapper van de goden, bekend om zijn gevleugelde sandalen?', 'sv': 'I grekisk mytologi, vilken gud ar gudarnas budbarare, kand for sina bevingade sandaler?', 'fr': "Dans la mythologie grecque, quel dieu est le messager des dieux, connu pour ses sandales ailees ?", 'da': 'I græsk mytologi, hvilken gud er gudernes budbringer, kendt for sine bevingede sandaler?'},
 'opt1': {'en': 'Hermes', 'de': 'Hermes', 'es': 'Hermes', 'nl': 'Hermes', 'sv': 'Hermes', 'fr': 'Hermes', 'da': 'Hermes'},
 'opt2': {'en': 'Apollo', 'de': 'Apollon', 'es': 'Apolo', 'nl': 'Apollo', 'sv': 'Apollon', 'fr': 'Apollon', 'da': 'Apollon'}},
]
