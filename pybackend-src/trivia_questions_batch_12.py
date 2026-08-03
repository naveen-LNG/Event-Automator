# -*- coding: utf-8 -*-
# Trivia Challenge batch #12. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which brown baked good is often eaten with jam and clotted cream in England?', 'de': 'Welches braune Gebaeck wird in England oft mit Marmelade und dicker Sahne gegessen?', 'es': '¿Que producto horneado marron se come a menudo con mermelada y nata cuajada en Inglaterra?', 'nl': 'Welk bruin gebak wordt in Engeland vaak gegeten met jam en dikke room?', 'sv': 'Vilket brunt bakverk ats ofta med sylt och tjock gradde i England?', 'fr': "Quelle patisserie brune se mange souvent avec de la confiture et de la creme epaisse en Angleterre ?", 'da': 'Hvilket brunt bagværk spises ofte med marmelade og tyk fløde i England?'},
 'opt1': {'en': 'Scone', 'de': 'Scone', 'es': 'Scone', 'nl': 'Scone', 'sv': 'Scones', 'fr': 'Scone', 'da': 'Scone'},
 'opt2': {'en': 'Muffin', 'de': 'Muffin', 'es': 'Muffin', 'nl': 'Muffin', 'sv': 'Muffins', 'fr': 'Muffin', 'da': 'Muffin'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Cairo?', 'de': 'In welchem Land liegt die antike Stadt Kairo?', 'es': '¿En que pais esta la antigua ciudad de El Cairo?', 'nl': 'In welk land ligt de oude stad Caïro?', 'sv': 'I vilket land ligger den antika staden Kairo?', 'fr': "Dans quel pays se trouve l'ancienne cite du Caire ?", 'da': 'Hvilket land huser den antikke by Kairo?'},
 'opt1': {'en': 'Egypt', 'de': 'Ägypten', 'es': 'Egipto', 'nl': 'Egypte', 'sv': 'Egypten', 'fr': "l'Egypte", 'da': 'Egypten'},
 'opt2': {'en': 'Libya', 'de': 'Libyen', 'es': 'Libia', 'nl': 'Libie', 'sv': 'Libyen', 'fr': 'la Libye', 'da': 'Libyen'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a fish who forgets things easily helping to find a lost son?', 'de': 'In welchem Filmklassiker hilft ein vergesslicher Fisch dabei, einen verlorenen Sohn zu finden?', 'es': '¿Que pelicula clasica presenta a un pez olvidadizo ayudando a encontrar a un hijo perdido?', 'nl': 'In welke klassieke film helpt een vergeetachtige vis om een verloren zoon te vinden?', 'sv': 'I vilken klassisk film hjalper en glomsk fisk till att hitta en forlorad son?', 'fr': "Quel film classique met en scene un poisson tres oublieux qui aide a retrouver un fils perdu ?", 'da': 'Hvilken klassisk film har en glemsom fisk, der hjælper med at finde en tabt søn?'},
 'opt1': {'en': 'Finding Nemo', 'de': 'Findet Nemo', 'es': 'Buscando a Nemo', 'nl': 'Finding Nemo', 'sv': 'Hitta Nemo', 'fr': 'Le Monde de Nemo', 'da': 'Find Nemo'},
 'opt2': {'en': 'Finding Dory', 'de': 'Findet Dorie', 'es': 'Buscando a Dory', 'nl': 'Finding Dory', 'sv': 'Hitta Doris', 'fr': 'Le Monde de Dory', 'da': 'Find Dory'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which pet is known for hopping in a cage and twitching its nose?', 'de': 'Welches Haustier huepft in einem Kaefig und zuckt mit der Nase?', 'es': '¿Que mascota es conocida por saltar en una jaula y mover la nariz?', 'nl': 'Welk huisdier huppelt in een kooi en trekt met zijn neus?', 'sv': 'Vilket sallskapsdjur hoppar i en bur och rycker pa nasan?', 'fr': "Quel animal de compagnie saute dans une cage et frotte son nez ?", 'da': 'Hvilket kæledyr hopper i et bur og rykker med næsen?'},
 'opt1': {'en': 'Rabbit', 'de': 'Kaninchen', 'es': 'Conejo', 'nl': 'Konijn', 'sv': 'Kanin', 'fr': 'Lapin', 'da': 'Kanin'},
 'opt2': {'en': 'Guinea pig', 'de': 'Meerschweinchen', 'es': 'Cobaya', 'nl': 'Cavia', 'sv': 'Marsvin', 'fr': "Cochon d'Inde", 'da': 'Marsvin'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which yellow fruit is used to make traditional banana splits?', 'de': 'Welche gelbe Frucht wird fuer traditionelle Bananensplit verwendet?', 'es': '¿Que fruta amarilla se usa para hacer el banana split tradicional?', 'nl': 'Welk geel fruit wordt gebruikt om traditionele banana splits te maken?', 'sv': 'Vilken gul frukt anvands for att gora traditionell banana split?', 'fr': "Quel fruit jaune est utilise pour faire le traditionnel banana split ?", 'da': 'Hvilken gul frugt bruges til at lave traditionel banana split?'},
 'opt1': {'en': 'Banana', 'de': 'Banane', 'es': 'Platano', 'nl': 'Banaan', 'sv': 'Banan', 'fr': 'Banane', 'da': 'Banan'},
 'opt2': {'en': 'Pineapple', 'de': 'Ananas', 'es': 'Piña', 'nl': 'Ananas', 'sv': 'Ananas', 'fr': 'Ananas', 'da': 'Ananas'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Petra and the capital Amman?', 'de': 'In welchem Land liegen die antike Stadt Petra und die Hauptstadt Amman?', 'es': '¿En que pais estan la antigua ciudad de Petra y la capital Amman?', 'nl': 'In welk land liggen de oude stad Petra en de hoofdstad Amman?', 'sv': 'I vilket land ligger den antika staden Petra och huvudstaden Amman?', 'fr': "Dans quel pays se trouvent l'ancienne cite de Petra et la capitale Amman ?", 'da': 'Hvilket land huser den antikke by Petra og hovedstaden Amman?'},
 'opt1': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'},
 'opt2': {'en': 'Lebanon', 'de': 'Libanon', 'es': 'Libano', 'nl': 'Libanon', 'sv': 'Libanon', 'fr': 'le Liban', 'da': 'Libanon'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. president is known for the Monroe Doctrine?', 'de': 'Welcher US-Praesident ist fuer die Monroe-Doktrin bekannt?', 'es': '¿Que presidente de EE.UU. es conocido por la Doctrina Monroe?', 'nl': 'Welke Amerikaanse president staat bekend om de Monroe-doctrine?', 'sv': 'Vilken amerikansk president ar kand for Monroedoktrinen?', 'fr': "Quel president americain est connu pour la doctrine Monroe ?", 'da': 'Hvilken amerikansk præsident er kendt for Monroe-doktrinen?'},
 'opt1': {'en': 'James Monroe', 'de': 'James Monroe', 'es': 'James Monroe', 'nl': 'James Monroe', 'sv': 'James Monroe', 'fr': 'James Monroe', 'da': 'James Monroe'},
 'opt2': {'en': 'John Quincy Adams', 'de': 'John Quincy Adams', 'es': 'John Quincy Adams', 'nl': 'John Quincy Adams', 'sv': 'John Quincy Adams', 'fr': 'John Quincy Adams', 'da': 'John Quincy Adams'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves painting on a canvas with oil or acrylic paints?', 'de': 'Bei welchem Handwerk wird auf Leinwand mit Ol- oder Acrylfarben gemalt?', 'es': '¿Que manualidad implica pintar sobre lienzo con oleos o acrilicos?', 'nl': 'Bij welk handwerk schilder je op canvas met olie- of acrylverf?', 'sv': 'Vilket hantverk innebar att mala pa duk med olje- eller akrylfarg?', 'fr': "Quel art manuel consiste a peindre sur toile a l'huile ou a l'acrylique ?", 'da': 'Hvilket håndværk involverer at male på lærred med olie- eller akrylmaling?'},
 'opt1': {'en': 'Painting', 'de': 'Malerei', 'es': 'Pintura', 'nl': 'Schilderen', 'sv': 'Malning', 'fr': 'La peinture', 'da': 'Maleri'},
 'opt2': {'en': 'Drawing', 'de': 'Zeichnen', 'es': 'Dibujo', 'nl': 'Tekenen', 'sv': 'Teckning', 'fr': 'Le dessin', 'da': 'Tegning'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which brass instrument is the smallest and highest-pitched in its family?', 'de': 'Welches Blechblasinstrument ist das kleinste und hoechste in seiner Familie?', 'es': '¿Que instrumento de metal es el mas pequeño y agudo de su familia?', 'nl': 'Welk koperblazersinstrument is het kleinst en hoogst in zijn familie?', 'sv': 'Vilket blasinstrument ar minst och har hogst tonlage i sin familj?', 'fr': "Quel instrument de cuivre est le plus petit et le plus aigu de sa famille ?", 'da': 'Hvilket messinginstrument er det mindste og højeste i sin familie?'},
 'opt1': {'en': 'Trumpet', 'de': 'Trompete', 'es': 'Trompeta', 'nl': 'Trompet', 'sv': 'Trumpet', 'fr': 'Trompette', 'da': 'Trompet'},
 'opt2': {'en': 'Trombone', 'de': 'Posaune', 'es': 'Trombon', 'nl': 'Trombone', 'sv': 'Trombon', 'fr': 'Trombone', 'da': 'Basun'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which natural event involves heavy rainfall causing rivers to overflow?', 'de': 'Welches Naturereignis beinhaltet starke Regenfaelle, die zum Ueberlaufen von Fluessen fuehren?', 'es': '¿Que evento natural implica fuertes lluvias que hacen desbordar los rios?', 'nl': 'Welke natuurlijke gebeurtenis omvat zware regenval waardoor rivieren overstromen?', 'sv': 'Vilken naturhandelse innebar kraftigt regn som far floder att svamma over?', 'fr': "Quel evenement naturel implique de fortes pluies faisant deborder les rivieres ?", 'da': 'Hvilken naturbegivenhed involverer kraftig regn, der får floder til at gå over?'},
 'opt1': {'en': 'Flood', 'de': 'Ueberschwemmung', 'es': 'Inundacion', 'nl': 'Overstroming', 'sv': 'Oversvamning', 'fr': 'Inondation', 'da': 'Oversvømmelse'},
 'opt2': {'en': 'Drought', 'de': 'Duerre', 'es': 'Sequia', 'nl': 'Droogte', 'sv': 'Torka', 'fr': 'Secheresse', 'da': 'Tørke'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a boy scout troop searching for pirate treasure?', 'de': 'In welchem Filmklassiker sucht eine Pfadfindergruppe nach Piratenschatz?', 'es': '¿Que pelicula clasica presenta a un grupo de exploradores buscando tesoro pirata?', 'nl': 'In welke klassieke film zoekt een groep padvinders naar piratenschatten?', 'sv': 'I vilken klassisk film letar en scouttrupp efter piratskatt?', 'fr': "Quel film classique met en scene une troupe de scouts a la recherche d'un tresor de pirate ?", 'da': 'Hvilken klassisk film har en spejdertrop, der søger efter piratskat?'},
 'opt1': {'en': 'The Goonies', 'de': 'Die Goonies', 'es': 'Los Goonies', 'nl': 'The Goonies', 'sv': 'Goonies', 'fr': 'Les Goonies', 'da': 'Goonierne'},
 'opt2': {'en': 'Treasure Island', 'de': 'Die Schatzinsel', 'es': 'La isla del tesoro', 'nl': 'Schateiland', 'sv': 'Skattkammarön', 'fr': "L'Ile au tresor", 'da': 'Skatteøen'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which U.S. state is known as the 'Beehive State'?", 'de': "Welcher US-Bundesstaat ist als 'Beehive State' bekannt?", 'es': "¿Que estado de EE.UU. es conocido como el 'Estado de la Colmena'?", 'nl': "Welke Amerikaanse staat staat bekend als de 'Beehive State'?", 'sv': "Vilken amerikansk delstat ar kand som 'Bikupstaten'?", 'fr': "Quel État americain est connu comme le 'Beehive State' (État de la ruche) ?", 'da': "Hvilken amerikansk delstat er kendt som 'Beehive State'?"},
 'opt1': {'en': 'Utah', 'de': 'Utah', 'es': 'Utah', 'nl': 'Utah', 'sv': 'Utah', 'fr': "l'Utah", 'da': 'Utah'},
 'opt2': {'en': 'Idaho', 'de': 'Idaho', 'es': 'Idaho', 'nl': 'Idaho', 'sv': 'Idaho', 'fr': "l'Idaho", 'da': 'Idaho'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which orange vegetable is used to make traditional pumpkin pie?', 'de': 'Welches orangefarbene Gemuese wird fuer traditionellen Kuerbiskuchen verwendet?', 'es': '¿Que verdura naranja se usa para hacer el pastel de calabaza tradicional?', 'nl': 'Welke oranje groente wordt gebruikt om traditionele pompoentaart te maken?', 'sv': 'Vilken orange gronsak anvands for att gora traditionell pumpapaj?', 'fr': "Quel legume orange est utilise pour faire la traditionnelle tarte a la citrouille ?", 'da': 'Hvilken orange grøntsag bruges til at lave traditionel græskartærte?'},
 'opt1': {'en': 'Pumpkin', 'de': 'Kuerbis', 'es': 'Calabaza', 'nl': 'Pompoen', 'sv': 'Pumpa', 'fr': 'La citrouille', 'da': 'Græskar'},
 'opt2': {'en': 'Sweet potato', 'de': 'Suesskartoffel', 'es': 'Camote', 'nl': 'Zoete aardappel', 'sv': 'Sotpotatis', 'fr': 'Patate douce', 'da': 'Sødkartoffel'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves lifting heavy weights overhead in competition?', 'de': 'Bei welcher Sportart werden schwere Gewichte im Wettbewerb ueber den Kopf gehoben?', 'es': '¿Que deporte consiste en levantar pesos pesados por encima de la cabeza en competicion?', 'nl': 'Bij welke sport til je zware gewichten boven het hoofd in competitie?', 'sv': 'Vilken sport innebar att lyfta tunga vikter over huvudet i tavling?', 'fr': "Quel sport consiste a soulever des poids lourds au-dessus de la tete en competition ?", 'da': 'Hvilken sport involverer at løfte tunge vægte over hovedet i konkurrence?'},
 'opt1': {'en': 'Weightlifting', 'de': 'Gewichtheben', 'es': 'Halterofilia', 'nl': 'Gewichtheffen', 'sv': 'Tyngdlyftning', 'fr': "L'halterophilie", 'da': 'Vægtløftning'},
 'opt2': {'en': 'Powerlifting', 'de': 'Kraftdreikampf', 'es': 'Powerlifting', 'nl': 'Powerlifting', 'sv': 'Styrkelyft', 'fr': 'La force athletique', 'da': 'Styrkeløft'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which large land mammal has antlers and is common in North American forests?', 'de': 'Welches grosse Landsaeugetier hat Geweihe und ist in nordamerikanischen Waeldern verbreitet?', 'es': '¿Que gran mamifero terrestre tiene cuernas y es comun en los bosques norteamericanos?', 'nl': 'Welk groot landzoogdier heeft een gewei en is gebruikelijk in Noord-Amerikaanse bossen?', 'sv': 'Vilket stort landdaggdjur har horn och ar vanligt i nordamerikanska skogar?', 'fr': "Quel grand mammifere terrestre a des bois et est courant dans les forets nord-americaines ?", 'da': 'Hvilket stort landpattedyr har gevir og er almindeligt i nordamerikanske skove?'},
 'opt1': {'en': 'Deer', 'de': 'Hirsch', 'es': 'Ciervo', 'nl': 'Hert', 'sv': 'Hjort', 'fr': 'Cerf', 'da': 'Hjort'},
 'opt2': {'en': 'Moose', 'de': 'Elch', 'es': 'Alce', 'nl': 'Eland', 'sv': 'Alg', 'fr': 'Elan', 'da': 'Elg'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which holiday celebrates the harvest with a cornucopia symbol in the fall?', 'de': 'Welcher Feiertag im Herbst feiert die Ernte mit dem Symbol des Füllhorns?', 'es': '¿Que festividad de otoño celebra la cosecha con un simbolo de cuerno de la abundancia?', 'nl': 'Welke feestdag in de herfst viert de oogst met een hoorn des overvloeds als symbool?', 'sv': 'Vilken hosthogtid firar skorden med en symbol av overflodshorn?', 'fr': "Quelle fete d'automne celebre la recolte avec le symbole de la corne d'abondance ?", 'da': 'Hvilken efterårshelligdag fejrer høsten med et overflødighedshorn som symbol?'},
 'opt1': {'en': 'Thanksgiving', 'de': 'Erntedankfest', 'es': 'Accion de Gracias', 'nl': 'Thanksgiving', 'sv': 'Thanksgiving', 'fr': 'Thanksgiving', 'da': 'Thanksgiving'},
 'opt2': {'en': 'Harvest Festival', 'de': 'Erntefest', 'es': 'Festival de la cosecha', 'nl': 'Oogstfeest', 'sv': 'Skordefest', 'fr': 'La fete des recoltes', 'da': 'Høstfest'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which continent is home to the Amazon River?', 'de': 'Auf welchem Kontinent liegt der Amazonas?', 'es': '¿En que continente esta el rio Amazonas?', 'nl': 'Op welk continent ligt de Amazonerivier?', 'sv': 'Pa vilken kontinent ligger Amazonfloden?', 'fr': "Sur quel continent se trouve le fleuve Amazone ?", 'da': 'Hvilket kontinent huser Amazonfloden?'},
 'opt1': {'en': 'South America', 'de': 'Suedamerika', 'es': 'Sudamerica', 'nl': 'Zuid-Amerika', 'sv': 'Sydamerika', 'fr': "l'Amerique du Sud", 'da': 'Sydamerika'},
 'opt2': {'en': 'North America', 'de': 'Nordamerika', 'es': 'Norteamerica', 'nl': 'Noord-Amerika', 'sv': 'Nordamerika', 'fr': "l'Amerique du Nord", 'da': 'Nordamerika'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which author wrote 'The Giving Tree'?", 'de': "Welcher Autor schrieb 'Der geheimnisvolle Baum'?", 'es': "¿Que autor escribio 'El arbol generoso'?", 'nl': "Welke auteur schreef 'De Geefboom'?", 'sv': "Vilken forfattare skrev 'Det givande tradet'?", 'fr': "Quel auteur a ecrit 'L'Arbre genereux' ?", 'da': "Hvilken forfatter skrev 'Det givende træ'?"},
 'opt1': {'en': 'Shel Silverstein', 'de': 'Shel Silverstein', 'es': 'Shel Silverstein', 'nl': 'Shel Silverstein', 'sv': 'Shel Silverstein', 'fr': 'Shel Silverstein', 'da': 'Shel Silverstein'},
 'opt2': {'en': 'Eric Carle', 'de': 'Eric Carle', 'es': 'Eric Carle', 'nl': 'Eric Carle', 'sv': 'Eric Carle', 'fr': 'Eric Carle', 'da': 'Eric Carle'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic toy is a doll that talks when you pull a string?', 'de': 'Welches klassische Spielzeug ist eine Puppe, die spricht, wenn man an einer Schnur zieht?', 'es': '¿Que juguete clasico es una muñeca que habla al tirar de un cordon?', 'nl': 'Welk klassiek speelgoed is een pop die praat als je aan een koordje trekt?', 'sv': 'Vilken klassisk leksak ar en docka som pratar nar man drar i ett snore?', 'fr': "Quel jouet classique est une poupee qui parle quand on tire sur une ficelle ?", 'da': 'Hvilket klassisk legetøj er en dukke, der taler, når man trækker i en snor?'},
 'opt1': {'en': 'Talking doll', 'de': 'Sprechende Puppe', 'es': 'Muñeca parlante', 'nl': 'Pratende pop', 'sv': 'Pratdocka', 'fr': 'Poupee parlante', 'da': 'Talende dukke'},
 'opt2': {'en': 'Chatty Cathy', 'de': 'Chatty Cathy', 'es': 'Chatty Cathy', 'nl': 'Chatty Cathy', 'sv': 'Chatty Cathy', 'fr': 'Chatty Cathy', 'da': 'Chatty Cathy'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which brown spice is used in chai tea alongside cardamom and ginger?', 'de': 'Welches braune Gewuerz wird in Chai-Tee neben Kardamom und Ingwer verwendet?', 'es': '¿Que especia marron se usa en el te chai junto al cardamomo y el jengibre?', 'nl': 'Welke bruine specerij wordt gebruikt in chaithee naast kardemom en gember?', 'sv': 'Vilken brun krydda anvands i chaite tillsammans med kardemumma och ingefara?', 'fr': "Quelle epice brune est utilisee dans le the chai aux cotes de la cardamome et du gingembre ?", 'da': 'Hvilket brunt krydderi bruges i chai-te sammen med kardemomme og ingefær?'},
 'opt1': {'en': 'Cinnamon', 'de': 'Zimt', 'es': 'Canela', 'nl': 'Kaneel', 'sv': 'Kanel', 'fr': 'La cannelle', 'da': 'Kanel'},
 'opt2': {'en': 'Clove', 'de': 'Gewuerznelke', 'es': 'Clavo', 'nl': 'Kruidnagel', 'sv': 'Kryddnejlika', 'fr': 'Le clou de girofle', 'da': 'Nellike'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. holiday commemorates fallen soldiers in late May?', 'de': 'Welcher US-Feiertag Ende Mai gedenkt gefallener Soldaten?', 'es': '¿Que festividad de EE.UU. a finales de mayo conmemora a los soldados caidos?', 'nl': 'Welke Amerikaanse feestdag eind mei herdenkt gesneuvelde soldaten?', 'sv': 'Vilken amerikansk hogtid i slutet av maj minns stupade soldater?', 'fr': "Quelle fete americaine de fin mai commemore les soldats tombes au combat ?", 'da': 'Hvilken amerikansk helligdag i slutningen af maj mindes faldne soldater?'},
 'opt1': {'en': 'Memorial Day', 'de': 'Memorial Day', 'es': 'Dia de los Caidos', 'nl': 'Memorial Day', 'sv': 'Memorial Day', 'fr': 'Le Memorial Day', 'da': 'Memorial Day'},
 'opt2': {'en': 'Veterans Day', 'de': 'Veterans Day', 'es': 'Dia de los Veteranos', 'nl': 'Veterans Day', 'sv': 'Veterandagen', 'fr': 'Le jour des Anciens combattants', 'da': 'Veterandag'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which decade is associated with the rise of country pop crossover artists like Shania Twain?', 'de': 'Mit welchem Jahrzehnt wird der Aufstieg von Country-Pop-Crossover-Kuenstlern wie Shania Twain assoziiert?', 'es': '¿Con que decada se asocia el auge de artistas de fusion country-pop como Shania Twain?', 'nl': 'Met welk decennium wordt de opkomst van country-popcrossoverartiesten zoals Shania Twain geassocieerd?', 'sv': 'Vilket artionde forknippas med framvaxten av country-pop-crossoverartister som Shania Twain?', 'fr': "A quelle decennie associe-t-on la montee d'artistes de fusion country-pop comme Shania Twain ?", 'da': 'Hvilket årti forbindes med fremkomsten af country-pop-crossoverkunstnere som Shania Twain?'},
 'opt1': {'en': '1990s', 'de': '1990er', 'es': 'los años 90', 'nl': 'jaren 90', 'sv': '1990-talet', 'fr': 'les annees 1990', 'da': '1990’erne'},
 'opt2': {'en': '1980s', 'de': '1980er', 'es': 'los años 80', 'nl': 'jaren 80', 'sv': '1980-talet', 'fr': 'les annees 1980', 'da': '1980’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Marrakesh?', 'de': 'In welchem Land liegt die antike Stadt Marrakesch?', 'es': '¿En que pais esta la antigua ciudad de Marrakech?', 'nl': 'In welk land ligt de oude stad Marrakech?', 'sv': 'I vilket land ligger den antika staden Marrakech?', 'fr': "Dans quel pays se trouve l'ancienne cite de Marrakech ?", 'da': 'Hvilket land huser den antikke by Marrakech?'},
 'opt1': {'en': 'Morocco', 'de': 'Marokko', 'es': 'Marruecos', 'nl': 'Marokko', 'sv': 'Marocko', 'fr': 'le Maroc', 'da': 'Marokko'},
 'opt2': {'en': 'Tunisia', 'de': 'Tunesien', 'es': 'Tunez', 'nl': 'Tunesie', 'sv': 'Tunisien', 'fr': 'la Tunisie', 'da': 'Tunesien'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which farm animal is known for gobbling and having a red wattle?', 'de': 'Welches Nutztier ist fuer sein Kollern und einen roten Kehllappen bekannt?', 'es': '¿Que animal de granja es conocido por su graznido y una barbilla roja?', 'nl': 'Welk boerderijdier staat bekend om zijn geluid en een rode lel?', 'sv': 'Vilket lantbruksdjur ar kant for sitt kluckande och en rod skäggflik?', 'fr': "Quel animal de ferme est connu pour son glouglou et son fanon rouge ?", 'da': 'Hvilket landbrugsdyr er kendt for sin skrattende lyd og en rød hagelap?'},
 'opt1': {'en': 'Turkey', 'de': 'Truthahn', 'es': 'Pavo', 'nl': 'Kalkoen', 'sv': 'Kalkon', 'fr': 'Dinde', 'da': 'Kalkun'},
 'opt2': {'en': 'Rooster', 'de': 'Hahn', 'es': 'Gallo', 'nl': 'Haan', 'sv': 'Tupp', 'fr': 'Coq', 'da': 'Hane'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which gemstone is the birthstone for those born in February?', 'de': 'Welcher Edelstein ist der Geburtsstein fuer im Februar Geborene?', 'es': '¿Que piedra preciosa es la piedra de nacimiento para los nacidos en febrero?', 'nl': 'Welke edelsteen is de geboortesteen voor mensen geboren in februari?', 'sv': 'Vilken adelsten ar fodelsestenen for de fodda i februari?', 'fr': "Quelle pierre precieuse est la pierre de naissance de fevrier ?", 'da': 'Hvilken ædelsten er fødselsstenen for dem, der er født i februar?'},
 'opt1': {'en': 'Amethyst', 'de': 'Amethyst', 'es': 'Amatista', 'nl': 'Amethist', 'sv': 'Ametist', 'fr': 'Amethyste', 'da': 'Ametyst'},
 'opt2': {'en': 'Garnet', 'de': 'Granat', 'es': 'Granate', 'nl': 'Granaat', 'sv': 'Granat', 'fr': 'Grenat', 'da': 'Granat'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic party game involves guessing charades acted out silently?', 'de': 'Bei welchem klassischen Partyspiel raet man stumm vorgefuehrte Scharaden?', 'es': '¿Que juego de fiesta clasico consiste en adivinar mimicas actuadas en silencio?', 'nl': 'Bij welk klassiek feestspel raad je stil uitgebeelde charades?', 'sv': 'Vilken klassisk festlek innebar att gissa charader som spelas upp tyst?', 'fr': "Quel jeu de fete classique consiste a deviner des mimes joues en silence ?", 'da': 'Hvilken klassisk selskabsleg involverer at gætte pantomimer, der spilles tavst?'},
 'opt1': {'en': 'Charades', 'de': 'Scharade', 'es': 'Mimica', 'nl': 'Charades', 'sv': 'Charader', 'fr': 'Les charades', 'da': 'Pantomime'},
 'opt2': {'en': 'Pictionary', 'de': 'Pictionary', 'es': 'Pictionary', 'nl': 'Pictionary', 'sv': 'Pictionary', 'fr': 'Pictionary', 'da': 'Pictionary'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which green herb is used to flavor traditional mint chocolate chip ice cream?', 'de': 'Welches gruene Kraut wird verwendet, um traditionelles Minz-Schokoladensplitter-Eis zu wuerzen?', 'es': '¿Que hierba verde se usa para dar sabor al helado tradicional de menta con chispas de chocolate?', 'nl': 'Welk groen kruid wordt gebruikt om traditioneel mint-chocoladeschilfersijs te smaken?', 'sv': 'Vilken gron ort anvands for att smaksatta traditionell mintchokladglass?', 'fr': "Quelle herbe verte est utilisee pour aromatiser la traditionnelle glace menthe-chocolat ?", 'da': 'Hvilken grøn urt bruges til at smage traditionel mynte-chokoladeis?'},
 'opt1': {'en': 'Mint', 'de': 'Minze', 'es': 'Menta', 'nl': 'Munt', 'sv': 'Mynta', 'fr': 'La menthe', 'da': 'Mynte'},
 'opt2': {'en': 'Basil', 'de': 'Basilikum', 'es': 'Albahaca', 'nl': 'Basilicum', 'sv': 'Basilika', 'fr': 'Le basilic', 'da': 'Basilikum'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which flower is traditionally associated with the birth month of June?', 'de': 'Welche Blume wird traditionell mit dem Geburtsmonat Juni assoziiert?', 'es': '¿Que flor se asocia tradicionalmente con el mes de nacimiento de junio?', 'nl': 'Welke bloem wordt traditioneel geassocieerd met de geboortemaand juni?', 'sv': 'Vilken blomma forknippas traditionellt med fodelsemanaden juni?', 'fr': "Quelle fleur est traditionnellement associee au mois de naissance de juin ?", 'da': 'Hvilken blomst forbindes traditionelt med fødselsmåneden juni?'},
 'opt1': {'en': 'Rose', 'de': 'Rose', 'es': 'Rosa', 'nl': 'Roos', 'sv': 'Ros', 'fr': 'Rose', 'da': 'Rose'},
 'opt2': {'en': 'Honeysuckle', 'de': 'Geissblatt', 'es': 'Madreselva', 'nl': 'Kamperfoelie', 'sv': 'Kaprifol', 'fr': 'Chevrefeuille', 'da': 'Kaprifolium'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Space Center Houston?', 'de': 'In welchem US-Bundesstaat liegt das Space Center Houston?', 'es': '¿En que estado de EE.UU. esta el Centro Espacial de Houston?', 'nl': 'In welke Amerikaanse staat ligt het Space Center Houston?', 'sv': 'I vilken amerikansk delstat ligger Space Center Houston?', 'fr': "Dans quel État americain se trouve le Space Center Houston ?", 'da': 'I hvilken amerikansk delstat ligger Space Center Houston?'},
 'opt1': {'en': 'Texas', 'de': 'Texas', 'es': 'Texas', 'nl': 'Texas', 'sv': 'Texas', 'fr': 'le Texas', 'da': 'Texas'},
 'opt2': {'en': 'Florida', 'de': 'Florida', 'es': 'Florida', 'nl': 'Florida', 'sv': 'Florida', 'fr': 'la Floride', 'da': 'Florida'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves carving intricate designs into pumpkins for autumn decorating?', 'de': 'Bei welchem Handwerk werden aufwaendige Muster in Kuerbisse geschnitzt fuer Herbstdekoration?', 'es': '¿Que manualidad implica tallar diseños intrincados en calabazas para decoracion otoñal?', 'nl': 'Bij welk handwerk snijd je ingewikkelde patronen in pompoenen voor herfstdecoratie?', 'sv': 'Vilket hantverk innebar att skara intrikata monster i pumpor for hostdekoration?', 'fr': "Quel art manuel consiste a sculpter des motifs complexes dans des citrouilles pour la decoration automnale ?", 'da': 'Hvilket håndværk involverer at udskære indviklede mønstre i græskar til efterårsdekoration?'},
 'opt1': {'en': 'Pumpkin carving', 'de': 'Kuerbisschnitzen', 'es': 'Talla de calabazas', 'nl': 'Pompoensnijden', 'sv': 'Pumpaskarning', 'fr': 'La sculpture de citrouille', 'da': 'Græskarudskæring'},
 'opt2': {'en': 'Gourd painting', 'de': 'Kuerbisbemalung', 'es': 'Pintura de calabazas', 'nl': 'Pompoenschilderen', 'sv': 'Pumpamalning', 'fr': 'La peinture de courges', 'da': 'Kalabasmaling'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves hitting a ball with a racquet against a single wall in a small court?', 'de': 'Bei welcher Sportart schlaegt man einen Ball mit einem Schlaeger gegen eine einzelne Wand in einem kleinen Platz?', 'es': '¿Que deporte consiste en golpear una pelota con una raqueta contra una sola pared en una cancha pequeña?', 'nl': 'Bij welke sport sla je een bal met een racket tegen een enkele muur in een kleine baan?', 'sv': 'Vilken sport innebar att slå en boll med en racket mot en enda vagg pa en liten bana?', 'fr': "Quel sport consiste a frapper une balle avec une raquette contre un seul mur dans un petit court ?", 'da': 'Hvilken sport involverer at slå en bold med en ketcher mod en enkelt væg i en lille bane?'},
 'opt1': {'en': 'Handball', 'de': 'Handball', 'es': 'Frontenis', 'nl': 'Muurkaatsen', 'sv': 'Handboll', 'fr': "Le handball (mur)", 'da': 'Håndbold'},
 'opt2': {'en': 'Racquetball', 'de': 'Racquetball', 'es': 'Racquetball', 'nl': 'Racquetball', 'sv': 'Racquetball', 'fr': 'Racquetball', 'da': 'Racquetball'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which item of clothing is worn around the neck to keep warm, often long and knitted?', 'de': 'Welches Kleidungsstueck wird um den Hals getragen, um warm zu bleiben, oft lang und gestrickt?', 'es': '¿Que prenda se lleva alrededor del cuello para mantener el calor, a menudo larga y tejida?', 'nl': 'Welk kledingstuk wordt om de nek gedragen om warm te blijven, vaak lang en gebreid?', 'sv': 'Vilket klädesplagg baers runt halsen for att halla varmen, ofta lang och stickad?', 'fr': "Quel vetement se porte autour du cou pour se tenir chaud, souvent long et tricote ?", 'da': 'Hvilken beklædningsgenstand bæres om halsen for at holde varmen, ofte lang og strikket?'},
 'opt1': {'en': 'Scarf', 'de': 'Schal (Halstuch)', 'es': 'Bufanda', 'nl': 'Sjaal', 'sv': 'Halsduk', 'fr': 'Echarpe', 'da': 'Tørklæde'},
 'opt2': {'en': 'Shawl', 'de': 'Schal', 'es': 'Chal', 'nl': 'Omslagdoek', 'sv': 'Sjal', 'fr': 'Chale', 'da': 'Sjal'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic cartoon features a blue hedgehog who runs fast to defeat a scientist?', 'de': 'In welchem klassischen Cartoon kommt ein blauer Igel vor, der schnell laeuft, um einen Wissenschaftler zu besiegen?', 'es': '¿Que dibujo animado clasico presenta a un erizo azul que corre rapido para derrotar a un cientifico?', 'nl': 'In welke klassieke tekenfilm komt een blauwe egel voor die snel rent om een wetenschapper te verslaan?', 'sv': 'I vilken klassisk tecknad film forekommer en bla igelkott som springer snabbt for att besegra en vetenskapsman?', 'fr': "Quel dessin anime classique met en scene un herisson bleu qui court vite pour vaincre un scientifique ?", 'da': 'Hvilken klassisk tegnefilm har et blåt pindsvin, der løber hurtigt for at besejre en videnskabsmand?'},
 'opt1': {'en': 'Sonic the Hedgehog', 'de': 'Sonic the Hedgehog', 'es': 'Sonic the Hedgehog', 'nl': 'Sonic the Hedgehog', 'sv': 'Sonic the Hedgehog', 'fr': 'Sonic the Hedgehog', 'da': 'Sonic the Hedgehog'},
 'opt2': {'en': 'Mega Man', 'de': 'Mega Man', 'es': 'Mega Man', 'nl': 'Mega Man', 'sv': 'Mega Man', 'fr': 'Mega Man', 'da': 'Mega Man'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Isfahan, known for its blue-tiled mosques?', 'de': 'In welchem Land liegt die antike Stadt Isfahan, bekannt fuer ihre blau gefliesten Moscheen?', 'es': '¿En que pais esta la antigua ciudad de Isfahan, conocida por sus mezquitas de azulejos azules?', 'nl': 'In welk land ligt de oude stad Isfahan, bekend om zijn blauw betegelde moskeeën?', 'sv': 'I vilket land ligger den antika staden Isfahan, kand for sina blakaklade moskeer?', 'fr': "Dans quel pays se trouve l'ancienne cite d'Ispahan, connue pour ses mosquees carrelees de bleu ?", 'da': 'Hvilket land huser den antikke by Isfahan, kendt for sine blåflisebelagte moskeer?'},
 'opt1': {'en': 'Iran', 'de': 'Iran', 'es': 'Iran', 'nl': 'Iran', 'sv': 'Iran', 'fr': "l'Iran", 'da': 'Iran'},
 'opt2': {'en': 'Iraq', 'de': 'Irak', 'es': 'Irak', 'nl': 'Irak', 'sv': 'Irak', 'fr': "l'Irak", 'da': 'Irak'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was known for her extensive support of NASA and the space program in public speeches?', 'de': 'Welche First Lady war bekannt fuer ihre umfangreiche Unterstuetzung der NASA und des Raumfahrtprogramms in oeffentlichen Reden?', 'es': '¿Que primera dama fue conocida por su extenso apoyo a la NASA y al programa espacial en discursos publicos?', 'nl': 'Welke First Lady stond bekend om haar uitgebreide steun aan NASA en het ruimtevaartprogramma in publieke toespraken?', 'sv': 'Vilken First Lady var kand for sitt omfattande stod for NASA och rymdprogrammet i offentliga tal?', 'fr': "Quelle Première dame etait connue pour son important soutien a la NASA et au programme spatial dans ses discours publics ?", 'da': 'Hvilken førstedame var kendt for sin omfattende støtte til NASA og rumprogrammet i offentlige taler?'},
 'opt1': {'en': 'Pat Nixon', 'de': 'Pat Nixon', 'es': 'Pat Nixon', 'nl': 'Pat Nixon', 'sv': 'Pat Nixon', 'fr': 'Pat Nixon', 'da': 'Pat Nixon'},
 'opt2': {'en': 'Betty Ford', 'de': 'Betty Ford', 'es': 'Betty Ford', 'nl': 'Betty Ford', 'sv': 'Betty Ford', 'fr': 'Betty Ford', 'da': 'Betty Ford'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Jamaican dish is a spiced meat pastry, similar to an empanada?', 'de': 'Welches jamaikanische Gericht ist ein gewuerztes Fleischgebaeck, aehnlich einer Empanada?', 'es': '¿Que plato jamaiquino es un pastel de carne especiado, similar a una empanada?', 'nl': 'Welk Jamaicaans gerecht is een gekruide vleespastei, vergelijkbaar met een empanada?', 'sv': 'Vilken jamaicansk ratt ar en kryddad kottpastej, liknande en empanada?', 'fr': "Quel plat jamaicain est un chausson de viande epicee, semblable a une empanada ?", 'da': 'Hvilken jamaicansk ret er en krydret kødpostej, ligner en empanada?'},
 'opt1': {'en': 'Jamaican patty', 'de': 'Jamaikanisches Pastete', 'es': 'Empanada jamaiquina', 'nl': 'Jamaicaans pasteitje', 'sv': 'Jamaicansk pastej', 'fr': 'Patty jamaicain', 'da': 'Jamaicansk patty'},
 'opt2': {'en': 'Beef pie', 'de': 'Rindfleischpastete', 'es': 'Pastel de carne', 'nl': 'Rundvleespastei', 'sv': 'Kottpaj', 'fr': 'La tourte au boeuf', 'da': 'Oksekødtærte'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which body system regulates temperature and includes skin, hair, and nails?', 'de': 'Welches Koerpersystem reguliert die Temperatur und umfasst Haut, Haare und Naegel?', 'es': '¿Que sistema corporal regula la temperatura e incluye piel, cabello y uñas?', 'nl': 'Welk lichaamssysteem reguleert temperatuur en omvat huid, haar en nagels?', 'sv': 'Vilket kroppssystem reglerar temperaturen och inkluderar hud, har och naglar?', 'fr': "Quel systeme corporel regule la temperature et comprend la peau, les cheveux et les ongles ?", 'da': 'Hvilket kropssystem regulerer temperaturen og omfatter hud, hår og negle?'},
 'opt1': {'en': 'Integumentary system', 'de': 'Integumentäres System', 'es': 'Sistema tegumentario', 'nl': 'Huidstelsel', 'sv': 'Hudsystemet', 'fr': 'Le systeme tegumentaire', 'da': 'Hudsystemet'},
 'opt2': {'en': 'Muscular system', 'de': 'Muskelsystem', 'es': 'Sistema muscular', 'nl': 'Spierstelsel', 'sv': 'Muskelsystemet', 'fr': 'Le systeme musculaire', 'da': 'Muskelsystemet'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which author wrote 'Anne of Green Gables'?", 'de': "Welche Autorin schrieb 'Anne auf Green Gables'?", 'es': "¿Que autora escribio 'Ana de las Tejas Verdes'?", 'nl': "Welke auteur schreef 'Anne of Green Gables'?", 'sv': "Vilken forfattare skrev 'Anne pa Gronkulla'?", 'fr': "Quelle autrice a ecrit 'Anne... la maison aux pignons verts' ?", 'da': "Hvilken forfatter skrev 'Anne fra Grønnebakken'?"},
 'opt1': {'en': 'L.M. Montgomery', 'de': 'L.M. Montgomery', 'es': 'L.M. Montgomery', 'nl': 'L.M. Montgomery', 'sv': 'L.M. Montgomery', 'fr': 'L.M. Montgomery', 'da': 'L.M. Montgomery'},
 'opt2': {'en': 'Laura Ingalls Wilder', 'de': 'Laura Ingalls Wilder', 'es': 'Laura Ingalls Wilder', 'nl': 'Laura Ingalls Wilder', 'sv': 'Laura Ingalls Wilder', 'fr': 'Laura Ingalls Wilder', 'da': 'Laura Ingalls Wilder'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Annie Wilkes in the film 'Misery'?", 'de': "Welche Schauspielerin spielte Annie Wilkes im Film 'Misery'?", 'es': "¿Que actriz interpreto a Annie Wilkes en la pelicula 'Misery'?", 'nl': "Welke actrice speelde Annie Wilkes in de film 'Misery'?", 'sv': "Vilken skadespelerska spelade Annie Wilkes i filmen 'Misery'?", 'fr': "Quelle actrice a joue Annie Wilkes dans le film 'Misery' ?", 'da': "Hvilken skuespiller spillede Annie Wilkes i filmen 'Misery'?"},
 'opt1': {'en': 'Kathy Bates', 'de': 'Kathy Bates', 'es': 'Kathy Bates', 'nl': 'Kathy Bates', 'sv': 'Kathy Bates', 'fr': 'Kathy Bates', 'da': 'Kathy Bates'},
 'opt2': {'en': 'Sissy Spacek', 'de': 'Sissy Spacek', 'es': 'Sissy Spacek', 'nl': 'Sissy Spacek', 'sv': 'Sissy Spacek', 'fr': 'Sissy Spacek', 'da': 'Sissy Spacek'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which singer is known for the album 'Jagged Little Pill'?", 'de': "Welche Saengerin ist fuer das Album 'Jagged Little Pill' bekannt?", 'es': "¿Que cantante es conocida por el album 'Jagged Little Pill'?", 'nl': "Welke zangeres staat bekend om het album 'Jagged Little Pill'?", 'sv': "Vilken sangerska ar kand for albumet 'Jagged Little Pill'?", 'fr': "Quelle chanteuse est connue pour l'album 'Jagged Little Pill' ?", 'da': "Hvilken sanger er kendt for albummet 'Jagged Little Pill'?"},
 'opt1': {'en': 'Alanis Morissette', 'de': 'Alanis Morissette', 'es': 'Alanis Morissette', 'nl': 'Alanis Morissette', 'sv': 'Alanis Morissette', 'fr': 'Alanis Morissette', 'da': 'Alanis Morissette'},
 'opt2': {'en': 'Sheryl Crow', 'de': 'Sheryl Crow', 'es': 'Sheryl Crow', 'nl': 'Sheryl Crow', 'sv': 'Sheryl Crow', 'fr': 'Sheryl Crow', 'da': 'Sheryl Crow'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Shiraz, known for gardens and poetry?', 'de': 'In welchem Land liegt die antike Stadt Schiraz, bekannt fuer Gaerten und Poesie?', 'es': '¿En que pais esta la antigua ciudad de Shiraz, conocida por sus jardines y poesia?', 'nl': 'In welk land ligt de oude stad Shiraz, bekend om tuinen en poëzie?', 'sv': 'I vilket land ligger den antika staden Shiraz, kand for tradgardar och poesi?', 'fr': "Dans quel pays se trouve l'ancienne cite de Chiraz, connue pour ses jardins et sa poesie ?", 'da': 'Hvilket land huser den antikke by Shiraz, kendt for haver og poesi?'},
 'opt1': {'en': 'Iran', 'de': 'Iran', 'es': 'Iran', 'nl': 'Iran', 'sv': 'Iran', 'fr': "l'Iran", 'da': 'Iran'},
 'opt2': {'en': 'Afghanistan', 'de': 'Afghanistan', 'es': 'Afganistan', 'nl': 'Afghanistan', 'sv': 'Afghanistan', 'fr': "l'Afghanistan", 'da': 'Afghanistan'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large aquatic reptile is the biggest of all living reptiles, found in Australia and Southeast Asia?', 'de': 'Welches grosse Wasserreptil ist das groesste aller lebenden Reptilien und lebt in Australien und Suedostasien?', 'es': '¿Que gran reptil acuatico es el mas grande de todos los reptiles vivos, encontrado en Australia y el sudeste asiatico?', 'nl': 'Welk groot waterreptiel is het grootste van alle levende reptielen, gevonden in Australie en Zuidoost-Azie?', 'sv': 'Vilket stort vattenlevande reptil ar det storsta av alla levande reptiler, finns i Australien och Sydostasien?', 'fr': "Quel grand reptile aquatique est le plus grand de tous les reptiles vivants, present en Australie et en Asie du Sud-Est ?", 'da': 'Hvilket stort vandkrybdyr er det største af alle levende krybdyr, fundet i Australien og Sydøstasien?'},
 'opt1': {'en': 'Saltwater crocodile', 'de': 'Leistenkrokodil', 'es': 'Cocodrilo marino', 'nl': 'Zoutwaterkrokodil', 'sv': 'Saltvattenskrokodil', 'fr': 'Crocodile marin', 'da': 'Saltvandskrokodille'},
 'opt2': {'en': 'Komodo dragon', 'de': 'Komodowaran', 'es': 'Dragon de Komodo', 'nl': 'Komodovaraan', 'sv': 'Komodovaran', 'fr': 'Dragon de Komodo', 'da': 'Komodovaran'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which South African dish is a spiced meat pie baked with a custard-like topping?', 'de': 'Welches suedafrikanische Gericht ist eine gewuerzte Fleischtorte, gebacken mit einem puddingartigen Belag?', 'es': '¿Que plato sudafricano es un pastel de carne especiada horneado con una cubierta similar a natillas?', 'nl': 'Welk Zuid-Afrikaans gerecht is een gekruide vleestaart gebakken met een custardachtige laag?', 'sv': 'Vilken sydafrikansk ratt ar en kryddad kottpaj bakad med en vaniljsasliknande topping?', 'fr': "Quel plat sud-africain est une tourte de viande epicee cuite avec une garniture semblable a de la creme anglaise ?", 'da': 'Hvilken sydafrikansk ret er en krydret kødtærte bagt med en custard-lignende topping?'},
 'opt1': {'en': 'Bobotie', 'de': 'Bobotie', 'es': 'Bobotie', 'nl': 'Bobotie', 'sv': 'Bobotie', 'fr': 'Bobotie', 'da': 'Bobotie'},
 'opt2': {'en': 'Bunny chow', 'de': 'Bunny Chow', 'es': 'Bunny chow', 'nl': 'Bunny chow', 'sv': 'Bunny chow', 'fr': 'Bunny chow', 'da': 'Bunny chow'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was the first to have earned a college degree before becoming First Lady?', 'de': 'Welche First Lady erwarb als erste einen Hochschulabschluss, bevor sie First Lady wurde?', 'es': '¿Que primera dama fue la primera en tener un titulo universitario antes de convertirse en primera dama?', 'nl': 'Welke First Lady behaalde als eerste een universitaire graad voordat ze First Lady werd?', 'sv': 'Vilken First Lady var forst med en universitetsexamen innan hon blev First Lady?', 'fr': "Quelle Première dame a ete la premiere a obtenir un diplome universitaire avant de devenir Première dame ?", 'da': 'Hvilken førstedame var den første til at have en universitetsgrad, før hun blev førstedame?'},
 'opt1': {'en': 'Lucy Hayes', 'de': 'Lucy Hayes', 'es': 'Lucy Hayes', 'nl': 'Lucy Hayes', 'sv': 'Lucy Hayes', 'fr': 'Lucy Hayes', 'da': 'Lucy Hayes'},
 'opt2': {'en': 'Frances Cleveland', 'de': 'Frances Cleveland', 'es': 'Frances Cleveland', 'nl': 'Frances Cleveland', 'sv': 'Frances Cleveland', 'fr': 'Frances Cleveland', 'da': 'Frances Cleveland'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which artist is known for large-scale color field paintings using rectangles of color?', 'de': 'Welcher Kuenstler ist fuer grossformatige Color-Field-Gemaelde mit farbigen Rechtecken bekannt?', 'es': '¿Que artista es conocido por pinturas de gran escala de campos de color usando rectangulos de color?', 'nl': 'Welke kunstenaar staat bekend om grootschalige color field-schilderijen met kleurrechthoeken?', 'sv': 'Vilken konstnar ar kand for storskaliga fargfaltsmalningar med fargrektanglar?', 'fr': "Quel artiste est connu pour ses peintures de champs de couleur a grande echelle utilisant des rectangles de couleur ?", 'da': 'Hvilken kunstner er kendt for store farvefeltmalerier med farverektangler?'},
 'opt1': {'en': 'Mark Rothko', 'de': 'Mark Rothko', 'es': 'Mark Rothko', 'nl': 'Mark Rothko', 'sv': 'Mark Rothko', 'fr': 'Mark Rothko', 'da': 'Mark Rothko'},
 'opt2': {'en': 'Barnett Newman', 'de': 'Barnett Newman', 'es': 'Barnett Newman', 'nl': 'Barnett Newman', 'sv': 'Barnett Newman', 'fr': 'Barnett Newman', 'da': 'Barnett Newman'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which flowering plant produces the spice vanilla from its seed pods?", 'de': 'Welche bluehende Pflanze produziert das Gewuerz Vanille aus ihren Samenkapseln?', 'es': '¿Que planta con flores produce la especia vainilla a partir de sus vainas de semillas?', 'nl': 'Welke bloeiende plant produceert de specerij vanille uit zijn zaadpeulen?', 'sv': 'Vilken blommande vaxt producerar kryddan vanilj fran sina fronkapslar?', 'fr': "Quelle plante a fleurs produit l'epice vanille a partir de ses gousses ?", 'da': 'Hvilken blomstrende plante producerer krydderiet vanilje fra sine frøkapsler?'},
 'opt1': {'en': 'Vanilla orchid', 'de': 'Vanille-Orchidee', 'es': 'Orquidea de vainilla', 'nl': 'Vanille-orchidee', 'sv': 'Vaniljorkide', 'fr': 'Orchidee vanille', 'da': 'Vaniljeorkidé'},
 'opt2': {'en': 'Vanilla bean tree', 'de': 'Vanilleschotenbaum', 'es': 'Arbol de vainilla', 'nl': 'Vanilleboon-boom', 'sv': 'Vaniljbönstrad', 'fr': 'Arbre a gousses de vanille', 'da': 'Vaniljebønnetræ'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which musical features the song 'Popular' sung by a witch before she turns green?", 'de': "In welchem Musical kommt das Lied 'Popular' vor, gesungen von einer Hexe, bevor sie gruen wird?", 'es': "¿Que musical presenta la cancion 'Popular' cantada por una bruja antes de volverse verde?", 'nl': "Welke musical bevat het lied 'Popular' gezongen door een heks voordat ze groen wordt?", 'sv': "Vilken musikal innehaller latten 'Popular' sjungen av en haxa innan hon blir gron?", 'fr': "Quelle comedie musicale comporte la chanson 'Popular' chantee par une sorciere avant qu'elle ne devienne verte ?", 'da': "Hvilken musical har sangen 'Popular' sunget af en heks, før hun bliver grøn?"},
 'opt1': {'en': 'Wicked', 'de': 'Wicked', 'es': 'Wicked', 'nl': 'Wicked', 'sv': 'Wicked', 'fr': 'Wicked', 'da': 'Wicked'},
 'opt2': {'en': 'The Wizard of Oz', 'de': 'Der Zauberer von Oz', 'es': 'El Mago de Oz', 'nl': 'The Wizard of Oz', 'sv': 'Trollkarlen fran Oz', 'fr': "Le Magicien d'Oz", 'da': 'Troldmanden fra Oz'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country has Reykjavik as its capital?', 'de': 'Welches Land hat Reykjavik als Hauptstadt?', 'es': '¿Que pais tiene a Reikiavik como su capital?', 'nl': 'Welk land heeft Reykjavik als hoofdstad?', 'sv': 'Vilket land har Reykjavik som huvudstad?', 'fr': "Quel pays a Reykjavik pour capitale ?", 'da': 'Hvilket land har Reykjavik som hovedstad?'},
 'opt1': {'en': 'Iceland', 'de': 'Island', 'es': 'Islandia', 'nl': 'IJsland', 'sv': 'Island', 'fr': "l'Islande", 'da': 'Island'},
 'opt2': {'en': 'Norway', 'de': 'Norwegen', 'es': 'Noruega', 'nl': 'Noorwegen', 'sv': 'Norge', 'fr': 'la Norvège', 'da': 'Norge'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which spice is used in traditional Ethiopian berbere seasoning blend?', 'de': 'Welches Gewuerz wird in traditioneller aethiopischer Berbere-Gewuerzmischung verwendet?', 'es': '¿Que especia se usa en la mezcla tradicional etiope de berbere?', 'nl': 'Welke specerij wordt gebruikt in traditionele Ethiopische berberekruidenmix?', 'sv': 'Vilken krydda anvands i traditionell etiopisk berbere-kryddblandning?', 'fr': "Quelle epice est utilisee dans le melange traditionnel ethiopien berbere ?", 'da': 'Hvilket krydderi bruges i traditionel etiopisk berbere-krydderiblanding?'},
 'opt1': {'en': 'Chili pepper', 'de': 'Chilischote', 'es': 'Chile', 'nl': 'Chilipeper', 'sv': 'Chilipeppar', 'fr': 'Piment', 'da': 'Chilipeber'},
 'opt2': {'en': 'Fenugreek', 'de': 'Bockshornklee', 'es': 'Fenogreco', 'nl': 'Fenegriek', 'sv': 'Bockhornsklover', 'fr': 'Fenugrec', 'da': 'Bukkehornsfrø'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which craft involves creating patterns by carving designs into linoleum or wood and printing them?', 'de': 'Bei welchem Handwerk werden Muster erzeugt, indem Motive in Linoleum oder Holz geschnitzt und gedruckt werden?', 'es': '¿Que manualidad implica crear patrones tallando diseños en linoleo o madera y estampandolos?', 'nl': 'Bij welk handwerk creëer je patronen door ontwerpen in linoleum of hout te snijden en af te drukken?', 'sv': 'Vilket hantverk innebar att skapa monster genom att skara motiv i linoleum eller trad och trycka dem?', 'fr': "Quel art manuel consiste a creer des motifs en sculptant des dessins dans du linoleum ou du bois et en les imprimant ?", 'da': 'Hvilket håndværk involverer at skabe mønstre ved at udskære motiver i linoleum eller træ og trykke dem?'},
 'opt1': {'en': 'Block printing', 'de': 'Blockdruck', 'es': 'Estampado en bloque', 'nl': 'Blokdruk', 'sv': 'Blocktryck', 'fr': 'La gravure sur bois (tampon)', 'da': 'Blokprint'},
 'opt2': {'en': 'Screen printing', 'de': 'Siebdruck', 'es': 'Serigrafia', 'nl': 'Zeefdruk', 'sv': 'Screentryck', 'fr': 'La serigraphie', 'da': 'Serigrafi'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which organ produces antibodies and filters blood, located near the stomach?', 'de': 'Welches Organ produziert Antikoerper und filtert Blut, in der Nahe des Magens gelegen?', 'es': '¿Que organo produce anticuerpos y filtra la sangre, ubicado cerca del estomago?', 'nl': 'Welk orgaan produceert antilichamen en filtert bloed, gelegen nabij de maag?', 'sv': 'Vilket organ producerar antikroppar och filtrerar blod, beläget nara magen?', 'fr': "Quel organe produit des anticorps et filtre le sang, situe pres de l'estomac ?", 'da': 'Hvilket organ producerer antistoffer og filtrerer blod, placeret nær maven?'},
 'opt1': {'en': 'Spleen', 'de': 'Milz', 'es': 'Bazo', 'nl': 'Milt', 'sv': 'Mjalte', 'fr': 'Rate', 'da': 'Milt'},
 'opt2': {'en': 'Pancreas', 'de': 'Bauchspeicheldruese', 'es': 'Pancreas', 'nl': 'Alvleesklier', 'sv': 'Bukspottkortel', 'fr': 'Pancreas', 'da': 'Bugspytkirtel'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'In archery, what is the term for the center of the target?', 'de': 'Wie nennt man beim Bogenschiessen die Mitte der Zielscheibe?', 'es': '¿En tiro con arco, como se llama el centro del blanco?', 'nl': 'Hoe heet bij boogschieten het midden van het doel?', 'sv': 'Vad kallas i bagskytte mitten av tavlan?', 'fr': "En tir a l'arc, comment appelle-t-on le centre de la cible ?", 'da': 'Hvad kaldes i bueskydning centrum af skiven?'},
 'opt1': {'en': 'Bullseye', 'de': 'Zentrum (Bullseye)', 'es': 'Diana (centro)', 'nl': 'Roos', 'sv': 'Pricken', 'fr': 'Le mille', 'da': 'Plet'},
 'opt2': {'en': 'Gold ring', 'de': 'Goldring', 'es': 'Anillo dorado', 'nl': 'Gouden ring', 'sv': 'Guldring', 'fr': 'Anneau doré', 'da': 'Guldring'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which Moroccan city is known as the 'Blue City' for its blue-painted buildings?", 'de': "Welche marokkanische Stadt ist als 'Blaue Stadt' bekannt fuer ihre blau bemalten Gebaeude?", 'es': "¿Que ciudad marroqui es conocida como la 'Ciudad Azul' por sus edificios pintados de azul?", 'nl': "Welke Marokkaanse stad staat bekend als de 'Blauwe Stad' vanwege zijn blauw geschilderde gebouwen?", 'sv': "Vilken marockansk stad ar kand som 'Bla staden' for sina bla malade byggnader?", 'fr': "Quelle ville marocaine est connue comme la 'Ville bleue' pour ses batiments peints en bleu ?", 'da': "Hvilken marokkansk by er kendt som 'Den blå by' for sine blåmalede bygninger?"},
 'opt1': {'en': 'Chefchaouen', 'de': 'Chefchaouen', 'es': 'Chefchaouen', 'nl': 'Chefchaouen', 'sv': 'Chefchaouen', 'fr': 'Chefchaouen', 'da': 'Chefchaouen'},
 'opt2': {'en': 'Fez', 'de': 'Fes', 'es': 'Fez', 'nl': 'Fez', 'sv': 'Fez', 'fr': 'Fes', 'da': 'Fez'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Miss Honey in the 1996 film 'Matilda'?", 'de': "Welche Schauspielerin spielte Miss Honey im Film 'Matilda' von 1996?", 'es': "¿Que actriz interpreto a la Señorita Honey en la pelicula de 1996 'Matilda'?", 'nl': "Welke actrice speelde Miss Honey in de film 'Matilda' uit 1996?", 'sv': "Vilken skadespelerska spelade Miss Honey i 1996 ars film 'Matilda'?", 'fr': "Quelle actrice a joue Miss Candy dans le film 'Matilda' de 1996 ?", 'da': "Hvilken skuespiller spillede Frøken Honey i 1996-filmen 'Matilda'?"},
 'opt1': {'en': 'Embeth Davidtz', 'de': 'Embeth Davidtz', 'es': 'Embeth Davidtz', 'nl': 'Embeth Davidtz', 'sv': 'Embeth Davidtz', 'fr': 'Embeth Davidtz', 'da': 'Embeth Davidtz'},
 'opt2': {'en': 'Rhea Perlman', 'de': 'Rhea Perlman', 'es': 'Rhea Perlman', 'nl': 'Rhea Perlman', 'sv': 'Rhea Perlman', 'fr': 'Rhea Perlman', 'da': 'Rhea Perlman'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which classic novel features the character Boo Radley?', 'de': 'In welchem Klassiker kommt die Figur Boo Radley vor?', 'es': '¿Que novela clasica presenta al personaje Boo Radley?', 'nl': 'In welke klassieke roman komt het personage Boo Radley voor?', 'sv': 'I vilken klassisk roman forekommer karaktaren Boo Radley?', 'fr': "Quel roman classique met en scene le personnage de Boo Radley ?", 'da': 'Hvilken klassisk roman har karakteren Boo Radley?'},
 'opt1': {'en': 'To Kill a Mockingbird', 'de': 'Wer die Nachtigall stoert', 'es': 'Matar a un ruiseñor', 'nl': 'Spaar de spotvogel', 'sv': 'Skri i natten', 'fr': "Ne tirez pas sur l'oiseau moqueur", 'da': 'Dræb ikke en sangfugl'},
 'opt2': {'en': 'The Catcher in the Rye', 'de': 'Der Faenger im Roggen', 'es': 'El guardian entre el centeno', 'nl': 'De vanger in het graan', 'sv': 'Ratten i vardagsrummet', 'fr': "L'Attrape-coeurs", 'da': 'Griberen i rugen'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which 1990s band is known for the song 'Wonderwall'?", 'de': "Welche Band der 1990er ist fuer das Lied 'Wonderwall' bekannt?", 'es': "¿Que banda de los 90 es conocida por la cancion 'Wonderwall'?", 'nl': "Welke band uit de jaren 90 staat bekend om het lied 'Wonderwall'?", 'sv': "Vilket 1990-talsband ar kant for latten 'Wonderwall'?", 'fr': "Quel groupe des annees 1990 est connu pour la chanson 'Wonderwall' ?", 'da': "Hvilket 1990’er-band er kendt for sangen 'Wonderwall'?"},
 'opt1': {'en': 'Oasis', 'de': 'Oasis', 'es': 'Oasis', 'nl': 'Oasis', 'sv': 'Oasis', 'fr': 'Oasis', 'da': 'Oasis'},
 'opt2': {'en': 'Blur', 'de': 'Blur', 'es': 'Blur', 'nl': 'Blur', 'sv': 'Blur', 'fr': 'Blur', 'da': 'Blur'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a savory bread pudding made from stale bread and eggs?', 'de': 'Welches suddstaatliche Gericht ist ein herzhafter Brotpudding aus altbackenem Brot und Eiern?', 'es': '¿Que plato sureño de EE.UU. es un pudin de pan salado hecho con pan duro y huevos?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een hartige broodpudding gemaakt van oudbakken brood en eieren?', 'sv': 'Vilken sydstatlig ratt ar en salt brodpudding gjord av gammalt brod och agg?', 'fr': "Quel plat du Sud des États-Unis est un pouding au pain sale fait de pain rassis et d'œufs ?", 'da': 'Hvilken sydstatsret er en salt brødbudding lavet af gammelt brød og æg?'},
 'opt1': {'en': 'Bread pudding', 'de': 'Brotpudding', 'es': 'Pudin de pan', 'nl': 'Broodpudding', 'sv': 'Brödpudding', 'fr': 'Pudding au pain', 'da': 'Brødbudding'},
 'opt2': {'en': 'French toast casserole', 'de': 'French-Toast-Auflauf', 'es': 'Cacerola de tostadas francesas', 'nl': 'Wentelteefjesschotel', 'sv': 'Fransk toast-gryta', 'fr': 'Gratin de pain perdu', 'da': 'Franskbrøds-gryderet'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which queen ruled Sweden and abdicated to convert to Catholicism in the 17th century?', 'de': 'Welche Koenigin regierte Schweden und dankte ab, um im 17. Jahrhundert zum Katholizismus zu konvertieren?', 'es': '¿Que reina goberno Suecia y abdico para convertirse al catolicismo en el siglo XVII?', 'nl': 'Welke koningin regeerde Zweden en trad af om zich in de 17e eeuw tot het katholicisme te bekeren?', 'sv': 'Vilken drottning styrde Sverige och abdikerade for att konvertera till katolicismen pa 1600-talet?', 'fr': "Quelle reine a regne sur la Suede et a abdique pour se convertir au catholicisme au 17e siecle ?", 'da': 'Hvilken dronning regerede Sverige og abdicerede for at konvertere til katolicismen i det 17. århundrede?'},
 'opt1': {'en': 'Queen Christina', 'de': 'Koenigin Christina', 'es': 'Reina Cristina', 'nl': 'Koningin Christina', 'sv': 'Drottning Kristina', 'fr': 'Reine Christine', 'da': 'Dronning Christina'},
 'opt2': {'en': 'Queen Margrethe', 'de': 'Koenigin Margrethe', 'es': 'Reina Margarita', 'nl': 'Koningin Margrethe', 'sv': 'Drottning Margrethe', 'fr': 'Reine Marguerite', 'da': 'Dronning Margrethe'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which small mammal is known for its distinctive black and white stripes, related to horses?', 'de': 'Welches kleine Saeugetier ist fuer seine markanten schwarz-weissen Streifen bekannt und mit Pferden verwandt?', 'es': '¿Que pequeño mamifero es conocido por sus distintivas rayas blancas y negras, emparentado con los caballos?', 'nl': 'Welk klein zoogdier staat bekend om zijn kenmerkende zwart-witte strepen en is verwant aan paarden?', 'sv': 'Vilket litet daggdjur ar kant for sina markanta svartvita ränder, slakt med hastar?', 'fr': "Quel petit mammifere est connu pour ses rayures noires et blanches distinctives, apparente aux chevaux ?", 'da': 'Hvilket lille pattedyr er kendt for sine karakteristiske sort-hvide striber, beslægtet med heste?'},
 'opt1': {'en': 'Zebra', 'de': 'Zebra', 'es': 'Cebra', 'nl': 'Zebra', 'sv': 'Zebra', 'fr': 'Zebre', 'da': 'Zebra'},
 'opt2': {'en': 'Okapi', 'de': 'Okapi', 'es': 'Okapi', 'nl': 'Okapi', 'sv': 'Okapi', 'fr': 'Okapi', 'da': 'Okapi'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Yazd, known for wind towers?', 'de': 'In welchem Land liegt die antike Stadt Yazd, bekannt fuer Windtuerme?', 'es': '¿En que pais esta la antigua ciudad de Yazd, conocida por sus torres de viento?', 'nl': 'In welk land ligt de oude stad Yazd, bekend om windtorens?', 'sv': 'I vilket land ligger den antika staden Yazd, kand for sina vindtorn?', 'fr': "Dans quel pays se trouve l'ancienne cite de Yazd, connue pour ses tours a vent ?", 'da': 'Hvilket land huser den antikke by Yazd, kendt for sine vindtårne?'},
 'opt1': {'en': 'Iran', 'de': 'Iran', 'es': 'Iran', 'nl': 'Iran', 'sv': 'Iran', 'fr': "l'Iran", 'da': 'Iran'},
 'opt2': {'en': 'Turkey', 'de': 'Tuerkei', 'es': 'Turquia', 'nl': 'Turkije', 'sv': 'Turkiet', 'fr': 'la Turquie', 'da': 'Tyrkiet'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of dress, fitted through the body with no waist seam, is a wardrobe staple?', 'de': 'Welche Art von Kleid, koerpernah ohne Taillennaht, ist ein Kleiderschrank-Grundstueck?', 'es': '¿Que tipo de vestido, ajustado sin costura en la cintura, es un basico del armario?', 'nl': 'Welk type jurk, aansluitend zonder taillenaad, is een basisstuk in de garderobe?', 'sv': 'Vilken typ av klanning, atsittande utan midjesöm, ar ett garderobsstapel?', 'fr': "Quel type de robe, ajustee sans couture a la taille, est un basique de la garde-robe ?", 'da': 'Hvilken type kjole, tætsiddende uden taljesøm, er en garderobefavorit?'},
 'opt1': {'en': 'Sheath dress', 'de': 'Etuikleid', 'es': 'Vestido tubo', 'nl': 'Kokerjurk', 'sv': 'Etuiklanning', 'fr': 'Robe fourreau', 'da': 'Etuikjole'},
 'opt2': {'en': 'Shift dress', 'de': 'Shift-Kleid', 'es': 'Vestido recto', 'nl': 'Shiftjurk', 'sv': 'Rak klanning', 'fr': 'Robe droite', 'da': 'Lige kjole'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which streaming series follows a group of women in a 1960s cooking show setting, based on a novel?', 'de': 'Welche Streaming-Serie folgt einer Gruppe von Frauen in einer Kochshow der 1960er Jahre, basierend auf einem Roman?', 'es': '¿Que serie de streaming sigue a un grupo de mujeres en un programa de cocina de los años 60, basada en una novela?', 'nl': 'Welke streamingserie volgt een groep vrouwen in een kookshow uit de jaren 60, gebaseerd op een roman?', 'sv': 'Vilken streamingserie foljer en grupp kvinnor i en matlagningsshow fran 1960-talet, baserad pa en roman?', 'fr': "Quelle serie en streaming suit un groupe de femmes dans une emission culinaire des annees 1960, adaptee d'un roman ?", 'da': 'Hvilken streamingserie følger en gruppe kvinder i et 1960’er-madlavningsprogram, baseret på en roman?'},
 'opt1': {'en': 'Lessons in Chemistry', 'de': 'Eine Frage der Chemie', 'es': 'Lecciones de quimica', 'nl': 'Lessons in Chemistry', 'sv': 'Lektioner i kemi', 'fr': 'Leçons de chimie', 'da': 'Lessons in Chemistry'},
 'opt2': {'en': 'Julia', 'de': 'Julia', 'es': 'Julia', 'nl': 'Julia', 'sv': 'Julia', 'fr': 'Julia', 'da': 'Julia'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which flower is known for its ruffled petals and comes in nearly every color except true blue?', 'de': 'Welche Blume ist fuer ihre gerueschten Blueten bekannt und kommt in fast jeder Farbe ausser echtem Blau vor?', 'es': '¿Que flor es conocida por sus petalos rizados y viene en casi todos los colores excepto el azul verdadero?', 'nl': 'Welke bloem staat bekend om zijn gekrulde bloemblaadjes en komt voor in bijna elke kleur behalve echt blauw?', 'sv': 'Vilken blomma ar kand for sina rufsiga kronblad och kommer i nastan alla farger utom aekta bla?', 'fr': "Quelle fleur est connue pour ses petales froisses et se decline dans presque toutes les couleurs sauf le vrai bleu ?", 'da': 'Hvilken blomst er kendt for sine krøllede kronblade og fås i næsten alle farver undtagen ægte blå?'},
 'opt1': {'en': 'Rose', 'de': 'Rose', 'es': 'Rosa', 'nl': 'Roos', 'sv': 'Ros', 'fr': 'Rose', 'da': 'Rose'},
 'opt2': {'en': 'Carnation', 'de': 'Nelke', 'es': 'Clavel', 'nl': 'Anjer', 'sv': 'Nejlika', 'fr': 'Oeillet', 'da': 'Nellike'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a fried dough pastry dusted with powdered sugar, popular in New Orleans?', 'de': 'Welches suddstaatliche Gericht ist ein frittiertes Teiggebaeck, bestaeubt mit Puderzucker, beliebt in New Orleans?', 'es': '¿Que plato sureño de EE.UU. es una masa frita espolvoreada con azucar glas, popular en Nueva Orleans?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een gefrituurd deeggebak bestrooid met poedersuiker, populair in New Orleans?', 'sv': 'Vilken sydstatlig ratt ar en friterad degvara pudrad med florsocker, popular i New Orleans?', 'fr': "Quel plat du Sud des États-Unis est une patisserie frite saupoudree de sucre glace, populaire a La Nouvelle-Orleans ?", 'da': 'Hvilken sydstatsret er en friturestegt dej drysset med flormelis, populær i New Orleans?'},
 'opt1': {'en': 'Beignet', 'de': 'Beignet', 'es': 'Beignet', 'nl': 'Beignet', 'sv': 'Beignet', 'fr': 'Beignet', 'da': 'Beignet'},
 'opt2': {'en': 'Zeppole', 'de': 'Zeppole', 'es': 'Zeppole', 'nl': 'Zeppole', 'sv': 'Zeppole', 'fr': 'Zeppole', 'da': 'Zeppole'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which vitamin is important for eye health and found in orange vegetables like carrots?', 'de': 'Welches Vitamin ist wichtig fuer die Augengesundheit und in orangefarbenem Gemuese wie Karotten enthalten?', 'es': '¿Que vitamina es importante para la salud ocular y se encuentra en verduras naranjas como las zanahorias?', 'nl': 'Welk vitamine is belangrijk voor de oogzondheid en te vinden in oranje groenten zoals wortels?', 'sv': 'Vilket vitamin ar viktigt for ogonhalsa och finns i orange gronsaker som morötter?', 'fr': "Quelle vitamine est importante pour la sante oculaire et se trouve dans les legumes orange comme les carottes ?", 'da': 'Hvilket vitamin er vigtigt for øjensundhed og findes i orange grøntsager som gulerødder?'},
 'opt1': {'en': 'Vitamin A', 'de': 'Vitamin A', 'es': 'Vitamina A', 'nl': 'Vitamine A', 'sv': 'Vitamin A', 'fr': 'Vitamine A', 'da': 'Vitamin A'},
 'opt2': {'en': 'Vitamin C', 'de': 'Vitamin C', 'es': 'Vitamina C', 'nl': 'Vitamine C', 'sv': 'Vitamin C', 'fr': 'Vitamine C', 'da': 'Vitamin C'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Tabriz?', 'de': 'In welchem Land liegt die antike Stadt Täbris?', 'es': '¿En que pais esta la antigua ciudad de Tabriz?', 'nl': 'In welk land ligt de oude stad Tabriz?', 'sv': 'I vilket land ligger den antika staden Tabriz?', 'fr': "Dans quel pays se trouve l'ancienne cite de Tabriz ?", 'da': 'Hvilket land huser den antikke by Tabriz?'},
 'opt1': {'en': 'Iran', 'de': 'Iran', 'es': 'Iran', 'nl': 'Iran', 'sv': 'Iran', 'fr': "l'Iran", 'da': 'Iran'},
 'opt2': {'en': 'Azerbaijan', 'de': 'Aserbaidschan', 'es': 'Azerbaiyan', 'nl': 'Azerbeidzjan', 'sv': 'Azerbajdzjan', 'fr': "l'Azerbaidjan", 'da': 'Aserbajdsjan'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was known for redecorating the White House in a colonial revival style?', 'de': 'Welche First Lady war bekannt fuer die Neugestaltung des Weissen Hauses im Colonial-Revival-Stil?', 'es': '¿Que primera dama fue conocida por redecorar la Casa Blanca en estilo colonial renacentista?', 'nl': 'Welke First Lady stond bekend om het herinrichten van het Witte Huis in koloniale revival-stijl?', 'sv': 'Vilken First Lady var kand for att inreda om Vita huset i kolonial vackelsestil?', 'fr': "Quelle Première dame etait connue pour avoir redecore la Maison-Blanche dans un style colonial renaissant ?", 'da': 'Hvilken førstedame var kendt for at genindrette Det Hvide Hus i koloniinspireret stil?'},
 'opt1': {'en': 'Grace Coolidge', 'de': 'Grace Coolidge', 'es': 'Grace Coolidge', 'nl': 'Grace Coolidge', 'sv': 'Grace Coolidge', 'fr': 'Grace Coolidge', 'da': 'Grace Coolidge'},
 'opt2': {'en': 'Lou Hoover', 'de': 'Lou Hoover', 'es': 'Lou Hoover', 'nl': 'Lou Hoover', 'sv': 'Lou Hoover', 'fr': 'Lou Hoover', 'da': 'Lou Hoover'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which composer wrote 'The Nutcracker' ballet?", 'de': "Welcher Komponist schrieb das Ballett 'Der Nussknacker'?", 'es': "¿Que compositor escribio el ballet 'El cascanueces'?", 'nl': "Welke componist schreef het ballet 'De Notenkraker'?", 'sv': "Vilken kompositor skrev baletten 'Nötknäpparen'?", 'fr': "Quel compositeur a ecrit le ballet 'Casse-Noisette' ?", 'da': "Hvilken komponist skrev balletten 'Nøddeknækkeren'?"},
 'opt1': {'en': 'Pyotr Ilyich Tchaikovsky', 'de': 'Pjotr Iljitsch Tschaikowski', 'es': 'Piotr Ilich Chaikovski', 'nl': 'Pjotr Iljitsj Tsjaikovski', 'sv': 'Pjotr Iljitj Tjajkovskij', 'fr': 'Piotr Ilitch Tchaikovski', 'da': 'Pjotr Iljitj Tjajkovskij'},
 'opt2': {'en': 'Sergei Prokofiev', 'de': 'Sergei Prokofjew', 'es': 'Sergei Prokofiev', 'nl': 'Sergej Prokofjev', 'sv': 'Sergej Prokofjev', 'fr': 'Sergueï Prokofiev', 'da': 'Sergej Prokofjev'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Andy in the film series 'The Devil Wears Prada'?", 'de': "Welche Schauspielerin spielte Andy in der Filmreihe 'Der Teufel traegt Prada'?", 'es': "¿Que actriz interpreto a Andy en la serie de peliculas 'El diablo viste a la moda'?", 'nl': "Welke actrice speelde Andy in de filmreeks 'The Devil Wears Prada'?", 'sv': "Vilken skadespelerska spelade Andy i filmserien 'Djavulen bar Prada'?", 'fr': "Quelle actrice a joue Andy dans la serie de films 'Le Diable s'habille en Prada' ?", 'da': "Hvilken skuespiller spillede Andy i filmserien 'Djævelen går i Prada'?"},
 'opt1': {'en': 'Anne Hathaway', 'de': 'Anne Hathaway', 'es': 'Anne Hathaway', 'nl': 'Anne Hathaway', 'sv': 'Anne Hathaway', 'fr': 'Anne Hathaway', 'da': 'Anne Hathaway'},
 'opt2': {'en': 'Emily Blunt', 'de': 'Emily Blunt', 'es': 'Emily Blunt', 'nl': 'Emily Blunt', 'sv': 'Emily Blunt', 'fr': 'Emily Blunt', 'da': 'Emily Blunt'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a savory casserole made with cheese grits and sausage for brunch?', 'de': 'Welches suddstaatliche Gericht ist ein herzhafter Auflauf aus Kaese-Grits und Wurst zum Brunch?', 'es': '¿Que plato sureño de EE.UU. es una cacerola salada hecha con grits con queso y salchicha para el brunch?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een hartige ovenschotel gemaakt met kaas-grits en worst voor brunch?', 'sv': 'Vilken sydstatlig ratt ar en salt gryta gjord med ostgrits och korv for brunch?', 'fr': "Quel plat du Sud des États-Unis est un gratin sale a base de grits au fromage et de saucisse pour le brunch ?", 'da': 'Hvilken sydstatsret er en salt gryderet lavet med ostegrits og pølse til brunch?'},
 'opt1': {'en': 'Sausage cheese grits casserole', 'de': 'Wurst-Kaese-Grits-Auflauf', 'es': 'Cacerola de grits con queso y salchicha', 'nl': 'Worst-kaas-gritsschotel', 'sv': 'Korv-ostgritsgryta', 'fr': 'Gratin de grits, fromage et saucisse', 'da': 'Pølse-ost-grits-gryderet'},
 'opt2': {'en': 'Breakfast casserole', 'de': 'Fruehstuecksauflauf', 'es': 'Cacerola de desayuno', 'nl': 'Ontbijtschotel', 'sv': 'Frukostgryta', 'fr': 'Gratin petit-dejeuner', 'da': 'Morgenmadsgryderet'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Berkshires mountain region?', 'de': 'In welchem US-Bundesstaat liegt die Bergregion Berkshires?', 'es': '¿En que estado de EE.UU. esta la region montañosa de los Berkshires?', 'nl': 'In welke Amerikaanse staat ligt de bergregio de Berkshires?', 'sv': 'I vilken amerikansk delstat ligger bergsregionen Berkshires?', 'fr': "Dans quel État americain se trouve la region montagneuse des Berkshires ?", 'da': 'I hvilken amerikansk delstat ligger bjergregionen Berkshires?'},
 'opt1': {'en': 'Massachusetts', 'de': 'Massachusetts', 'es': 'Massachusetts', 'nl': 'Massachusetts', 'sv': 'Massachusetts', 'fr': 'le Massachusetts', 'da': 'Massachusetts'},
 'opt2': {'en': 'Vermont', 'de': 'Vermont', 'es': 'Vermont', 'nl': 'Vermont', 'sv': 'Vermont', 'fr': 'le Vermont', 'da': 'Vermont'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large bird is flightless, native to Africa, and the fastest running bird?', 'de': 'Welcher grosse Vogel ist flugunfaehig, in Afrika heimisch und der schnellste laufende Vogel?', 'es': '¿Que gran ave no vuela, es nativa de Africa y es el ave corredora mas rapida?', 'nl': 'Welke grote vogel is vliegloos, inheems in Afrika en de snelste rennende vogel?', 'sv': 'Vilken stor fagel ar flygofor, inhemsk i Afrika och den snabbaste springande fageln?', 'fr': "Quel grand oiseau est incapable de voler, originaire d'Afrique, et l'oiseau coureur le plus rapide ?", 'da': 'Hvilken stor fugl er flyveudygtig, hjemmehørende i Afrika, og den hurtigst løbende fugl?'},
 'opt1': {'en': 'Ostrich', 'de': 'Strauss', 'es': 'Avestruz', 'nl': 'Struisvogel', 'sv': 'Struts', 'fr': 'Autruche', 'da': 'Struds'},
 'opt2': {'en': 'Rhea', 'de': 'Nandu', 'es': 'Ñandu', 'nl': 'Nandoe', 'sv': 'Nandu', 'fr': 'Nandou', 'da': 'Nandu'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which Persian New Year tradition involves a table setting with seven symbolic items starting with 'S'?", 'de': "Welche persische Neujahrstradition beinhaltet eine Tischdekoration mit sieben symbolischen Gegenstaenden, die mit 'S' beginnen?", 'es': "¿Que tradicion del Año Nuevo persa implica una mesa con siete elementos simbolicos que empiezan con la letra 'S'?", 'nl': "Welke Perzische nieuwjaarstraditie omvat een tafelschikking met zeven symbolische items die beginnen met 'S'?", 'sv': "Vilken persisk nyarstradition innebar en bordsuppsattning med sju symboliska foremal som borjar pa 'S'?", 'fr': "Quelle tradition du Nouvel An persan implique une table avec sept elements symboliques commençant par 'S' ?", 'da': "Hvilken persisk nytårstradition involverer en borddækning med syv symbolske ting, der begynder med 'S'?"},
 'opt1': {'en': 'Nowruz (Haft-sin)', 'de': 'Nowruz (Haft-Sin)', 'es': 'Nowruz (Haft-sin)', 'nl': 'Nowruz (Haft-sin)', 'sv': 'Nowruz (Haft-sin)', 'fr': 'Norouz (Haft-sin)', 'da': 'Nowruz (Haft-sin)'},
 'opt2': {'en': 'Ramadan', 'de': 'Ramadan', 'es': 'Ramadan', 'nl': 'Ramadan', 'sv': 'Ramadan', 'fr': 'Le ramadan', 'da': 'Ramadan'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which author wrote 'The Poisonwood Bible'?", 'de': "Welche Autorin schrieb 'Die Giftholzbibel'?", 'es': "¿Que autora escribio 'Los ojos en la noche' (The Poisonwood Bible)?", 'nl': "Welke auteur schreef 'The Poisonwood Bible'?", 'sv': "Vilken forfattare skrev 'Giftbibeln'?", 'fr': "Quelle autrice a ecrit 'Les Yeux dans l'arbre' (The Poisonwood Bible) ?", 'da': "Hvilken forfatter skrev 'The Poisonwood Bible'?"},
 'opt1': {'en': 'Barbara Kingsolver', 'de': 'Barbara Kingsolver', 'es': 'Barbara Kingsolver', 'nl': 'Barbara Kingsolver', 'sv': 'Barbara Kingsolver', 'fr': 'Barbara Kingsolver', 'da': 'Barbara Kingsolver'},
 'opt2': {'en': 'Ann Patchett', 'de': 'Ann Patchett', 'es': 'Ann Patchett', 'nl': 'Ann Patchett', 'sv': 'Ann Patchett', 'fr': 'Ann Patchett', 'da': 'Ann Patchett'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Susa?', 'de': 'In welchem Land liegt die antike Stadt Susa?', 'es': '¿En que pais esta la antigua ciudad de Susa?', 'nl': 'In welk land ligt de oude stad Susa?', 'sv': 'I vilket land ligger den antika staden Susa?', 'fr': "Dans quel pays se trouve l'ancienne cite de Suse ?", 'da': 'Hvilket land huser den antikke by Susa?'},
 'opt1': {'en': 'Iran', 'de': 'Iran', 'es': 'Iran', 'nl': 'Iran', 'sv': 'Iran', 'fr': "l'Iran", 'da': 'Iran'},
 'opt2': {'en': 'Iraq', 'de': 'Irak', 'es': 'Irak', 'nl': 'Irak', 'sv': 'Irak', 'fr': "l'Irak", 'da': 'Irak'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which instrument is central to Persian classical music, a long-necked lute?', 'de': 'Welches Instrument ist zentral fuer persische klassische Musik, eine langhalsige Laute?', 'es': '¿Que instrumento es central en la musica clasica persa, un laud de mastil largo?', 'nl': 'Welk instrument is centraal in Perzische klassieke muziek, een langhalzige luit?', 'sv': 'Vilket instrument ar centralt i persisk klassisk musik, en langhalsad luta?', 'fr': "Quel instrument est central dans la musique classique persane, un luth au long manche ?", 'da': 'Hvilket instrument er centralt i persisk klassisk musik, en langhalset lut?'},
 'opt1': {'en': 'Tar', 'de': 'Tar', 'es': 'Tar', 'nl': 'Tar', 'sv': 'Tar', 'fr': 'Tar', 'da': 'Tar'},
 'opt2': {'en': 'Oud', 'de': 'Oud', 'es': 'Oud', 'nl': 'Oud', 'sv': 'Oud', 'fr': 'Oud', 'da': 'Oud'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a slow-cooked pork shoulder, often served pulled with sauce?', 'de': 'Welches suddstaatliche Gericht ist eine langsam gegarte Schweineschulter, oft zerzupft mit Sauce serviert?', 'es': '¿Que plato sureño de EE.UU. es una paletilla de cerdo cocida a fuego lento, a menudo servida desmenuzada con salsa?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een langzaam gegaarde varkensschouder, vaak geserveerd getrokken met saus?', 'sv': 'Vilken sydstatlig ratt ar en langsamt tillagad flaskbog, ofta serverad rivet med sas?', 'fr': "Quel plat du Sud des États-Unis est une epaule de porc cuite lentement, souvent servie effilochee avec de la sauce ?", 'da': 'Hvilken sydstatsret er en langsomt tilberedt svinebov, ofte serveret trukket med sovs?'},
 'opt1': {'en': 'Pulled pork', 'de': 'Pulled Pork', 'es': 'Cerdo desmenuzado', 'nl': 'Pulled pork', 'sv': 'Pulled pork', 'fr': 'Porc effiloche', 'da': 'Pulled pork'},
 'opt2': {'en': 'Pork ribs', 'de': 'Schweinerippchen', 'es': 'Costillas de cerdo', 'nl': 'Varkensribbetjes', 'sv': 'Flaskribbor', 'fr': 'Cotes de porc', 'da': 'Svineribben'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which mineral, found in bananas and potatoes, helps prevent muscle cramps?', 'de': 'Welches Mineral, in Bananen und Kartoffeln enthalten, hilft, Muskelkraempfen vorzubeugen?', 'es': '¿Que mineral, presente en platanos y papas, ayuda a prevenir calambres musculares?', 'nl': 'Welk mineraal, aanwezig in bananen en aardappelen, helpt spierkrampen te voorkomen?', 'sv': 'Vilket mineral, som finns i bananer och potatis, hjalper till att forebygga muskelkramper?', 'fr': "Quel mineral, present dans les bananes et les pommes de terre, aide a prevenir les crampes musculaires ?", 'da': 'Hvilket mineral, fundet i bananer og kartofler, hjælper med at forhindre muskelkramper?'},
 'opt1': {'en': 'Potassium', 'de': 'Kalium', 'es': 'Potasio', 'nl': 'Kalium', 'sv': 'Kalium', 'fr': 'Potassium', 'da': 'Kalium'},
 'opt2': {'en': 'Sodium', 'de': 'Natrium', 'es': 'Sodio', 'nl': 'Natrium', 'sv': 'Natrium', 'fr': 'Sodium', 'da': 'Natrium'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which film has Reese Witherspoon playing a woman who hikes the Pacific Crest Trail alone?', 'de': 'In welchem Film spielt Reese Witherspoon eine Frau, die allein den Pacific Crest Trail wandert?', 'es': '¿En que pelicula Reese Witherspoon interpreta a una mujer que recorre sola el Sendero Pacific Crest?', 'nl': 'In welke film speelt Reese Witherspoon een vrouw die alleen de Pacific Crest Trail wandelt?', 'sv': 'I vilken film spelar Reese Witherspoon en kvinna som vandrar Pacific Crest Trail ensam?', 'fr': "Dans quel film Reese Witherspoon joue-t-elle une femme qui parcourt seule le Pacific Crest Trail ?", 'da': 'I hvilken film spiller Reese Witherspoon en kvinde, der vandrer Pacific Crest Trail alene?'},
 'opt1': {'en': 'Wild', 'de': 'Der große Trip - Wild', 'es': 'Salvaje', 'nl': 'Wild', 'sv': 'Wild', 'fr': 'Wild', 'da': 'Wild'},
 'opt2': {'en': 'Eat Pray Love', 'de': 'Eat Pray Love', 'es': 'Comer, rezar, amar', 'nl': 'Eat Pray Love', 'sv': 'Ät, be, alska', 'fr': 'Mange, prie, aime', 'da': 'Eat, Pray, Love'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Kashan, known for rosewater production?', 'de': 'In welchem Land liegt die antike Stadt Kaschan, bekannt fuer Rosenwasserproduktion?', 'es': '¿En que pais esta la antigua ciudad de Kashan, conocida por la produccion de agua de rosas?', 'nl': 'In welk land ligt de oude stad Kashan, bekend om rozenwaterproductie?', 'sv': 'I vilket land ligger den antika staden Kashan, kand for rosvattenproduktion?', 'fr': "Dans quel pays se trouve l'ancienne cite de Kashan, connue pour la production d'eau de rose ?", 'da': 'Hvilket land huser den antikke by Kashan, kendt for rosenvandsproduktion?'},
 'opt1': {'en': 'Iran', 'de': 'Iran', 'es': 'Iran', 'nl': 'Iran', 'sv': 'Iran', 'fr': "l'Iran", 'da': 'Iran'},
 'opt2': {'en': 'Iraq', 'de': 'Irak', 'es': 'Irak', 'nl': 'Irak', 'sv': 'Irak', 'fr': "l'Irak", 'da': 'Irak'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of glasses became a fashion statement popularized by aviators in the military?', 'de': 'Welche Art von Brille wurde zum Modestatement, populaer gemacht von Militaerfliegern?', 'es': '¿Que tipo de gafas se convirtieron en una declaracion de moda popularizada por aviadores militares?', 'nl': 'Welk type bril werd een modestatement gepopulariseerd door piloten in het leger?', 'sv': 'Vilken typ av glasogon blev en modemarkering popularaiserad av flygare i militaren?', 'fr': "Quel type de lunettes est devenu une declaration de mode popularisee par les aviateurs militaires ?", 'da': 'Hvilken type briller blev en modeerklæring populariseret af flyvere i militæret?'},
 'opt1': {'en': 'Aviator sunglasses', 'de': 'Pilotenbrille', 'es': 'Gafas de aviador', 'nl': 'Aviator-zonnebril', 'sv': 'Pilotglasogon', 'fr': 'Lunettes aviateur', 'da': 'Pilotbriller'},
 'opt2': {'en': 'Wayfarer sunglasses', 'de': 'Wayfarer-Brille', 'es': 'Gafas Wayfarer', 'nl': 'Wayfarer-zonnebril', 'sv': 'Wayfarer-solglasogon', 'fr': 'Lunettes Wayfarer', 'da': 'Wayfarer-solbriller'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large marine fish is known for its size, filter feeding, and gentle nature?', 'de': 'Welcher grosse Meeresfisch ist fuer seine Groesse, seine Filterernaehrung und sein sanftes Wesen bekannt?', 'es': '¿Que gran pez marino es conocido por su tamaño, alimentacion por filtracion y naturaleza gentil?', 'nl': 'Welke grote zeevis staat bekend om zijn grootte, filtervoeding en zachtaardige aard?', 'sv': 'Vilken stor havsfisk ar kand for sin storlek, filtermatning och milda natur?', 'fr': "Quel grand poisson marin est connu pour sa taille, son alimentation par filtration et sa nature douce ?", 'da': 'Hvilken stor havfisk er kendt for sin størrelse, filterfødning og milde natur?'},
 'opt1': {'en': 'Whale shark', 'de': 'Walhai', 'es': 'Tiburon ballena', 'nl': 'Walvishaai', 'sv': 'Valhaj', 'fr': 'Requin-baleine', 'da': 'Hvalhaj'},
 'opt2': {'en': 'Basking shark', 'de': 'Riesenhai', 'es': 'Tiburon peregrino', 'nl': 'Reuzenhaai', 'sv': 'Brugd', 'fr': 'Requin pelerin', 'da': 'Brugde'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was known for hosting weekly radio addresses and press conferences for women only?', 'de': 'Welche First Lady war bekannt fuer woechentliche Radioansprachen und Pressekonferenzen nur fuer Frauen?', 'es': '¿Que primera dama fue conocida por sus discursos radiales semanales y conferencias de prensa solo para mujeres?', 'nl': 'Welke First Lady stond bekend om wekelijkse radiotoespraken en persconferenties alleen voor vrouwen?', 'sv': 'Vilken First Lady var kand for veckovisa radiotal och presskonferenser bara for kvinnor?', 'fr': "Quelle Première dame etait connue pour ses allocutions radio hebdomadaires et ses conferences de presse reservees aux femmes ?", 'da': 'Hvilken førstedame var kendt for ugentlige radiotaler og pressekonferencer kun for kvinder?'},
 'opt1': {'en': 'Eleanor Roosevelt', 'de': 'Eleanor Roosevelt', 'es': 'Eleanor Roosevelt', 'nl': 'Eleanor Roosevelt', 'sv': 'Eleanor Roosevelt', 'fr': 'Eleanor Roosevelt', 'da': 'Eleanor Roosevelt'},
 'opt2': {'en': 'Lou Hoover', 'de': 'Lou Hoover', 'es': 'Lou Hoover', 'nl': 'Lou Hoover', 'sv': 'Lou Hoover', 'fr': 'Lou Hoover', 'da': 'Lou Hoover'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which First Lady was the first to be a divorcee before marrying the president?', 'de': 'Welche First Lady war die erste geschiedene Frau, bevor sie den Praesidenten heiratete?', 'es': '¿Que primera dama fue la primera divorciada antes de casarse con el presidente?', 'nl': 'Welke First Lady was de eerste gescheiden vrouw voordat ze met de president trouwde?', 'sv': 'Vilken First Lady var forst med att vara skild innan hon gifte sig med presidenten?', 'fr': "Quelle Première dame a ete la premiere divorcee avant d'epouser le president ?", 'da': 'Hvilken førstedame var den første til at være skilt, før hun giftede sig med præsidenten?'},
 'opt1': {'en': 'Rachel Jackson', 'de': 'Rachel Jackson', 'es': 'Rachel Jackson', 'nl': 'Rachel Jackson', 'sv': 'Rachel Jackson', 'fr': 'Rachel Jackson', 'da': 'Rachel Jackson'},
 'opt2': {'en': 'Florence Harding', 'de': 'Florence Harding', 'es': 'Florence Harding', 'nl': 'Florence Harding', 'sv': 'Florence Harding', 'fr': 'Florence Harding', 'da': 'Florence Harding'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Turkish author won the Nobel Prize in Literature in 2006?', 'de': 'Welcher tuerkische Autor gewann 2006 den Nobelpreis fuer Literatur?', 'es': '¿Que autor turco gano el Premio Nobel de Literatura en 2006?', 'nl': 'Welke Turkse auteur won in 2006 de Nobelprijs voor Literatuur?', 'sv': 'Vilken turkisk forfattare vann Nobelpriset i litteratur 2006?', 'fr': "Quel auteur turc a remporte le prix Nobel de litterature en 2006 ?", 'da': 'Hvilken tyrkisk forfatter vandt Nobelprisen i litteratur i 2006?'},
 'opt1': {'en': 'Orhan Pamuk', 'de': 'Orhan Pamuk', 'es': 'Orhan Pamuk', 'nl': 'Orhan Pamuk', 'sv': 'Orhan Pamuk', 'fr': 'Orhan Pamuk', 'da': 'Orhan Pamuk'},
 'opt2': {'en': 'Yashar Kemal', 'de': 'Yaşar Kemal', 'es': 'Yashar Kemal', 'nl': 'Yasar Kemal', 'sv': 'Yasar Kemal', 'fr': 'Yachar Kemal', 'da': 'Yasar Kemal'}},

{'category': 'Science', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which vitamin, essential for DNA synthesis, is often supplemented during pregnancy?', 'de': 'Welches Vitamin, essentiell fuer die DNA-Synthese, wird waehrend der Schwangerschaft oft zusaetzlich eingenommen?', 'es': '¿Que vitamina, esencial para la sintesis del ADN, se suplementa a menudo durante el embarazo?', 'nl': 'Welk vitamine, essentieel voor DNA-synthese, wordt vaak aangevuld tijdens de zwangerschap?', 'sv': 'Vilket vitamin, viktigt for DNA-syntes, tilläggs ofta under graviditeten?', 'fr': "Quelle vitamine, essentielle a la synthese de l'ADN, est souvent supplementee pendant la grossesse ?", 'da': 'Hvilket vitamin, essentielt for DNA-syntese, tilskuds ofte under graviditet?'},
 'opt1': {'en': 'Folic acid (Vitamin B9)', 'de': 'Folsaeure (Vitamin B9)', 'es': 'Acido folico (Vitamina B9)', 'nl': 'Foliumzuur (Vitamine B9)', 'sv': 'Folsyra (Vitamin B9)', 'fr': 'Acide folique (Vitamine B9)', 'da': 'Folinsyre (Vitamin B9)'},
 'opt2': {'en': 'Vitamin B6', 'de': 'Vitamin B6', 'es': 'Vitamina B6', 'nl': 'Vitamine B6', 'sv': 'Vitamin B6', 'fr': 'Vitamine B6', 'da': 'Vitamin B6'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Austrian artist created decorative gold-leaf paintings including 'The Kiss'?", 'de': "Welcher oesterreichische Kuenstler schuf dekorative Goldblatt-Gemaelde, darunter 'Der Kuss'?", 'es': "¿Que artista austriaco creo pinturas decorativas de pan de oro incluyendo 'El beso'?", 'nl': "Welke Oostenrijkse kunstenaar creëerde decoratieve bladgoud-schilderijen waaronder 'De Kus'?", 'sv': "Vilken osterrikisk konstnar skapade dekorativa bladguldmalningar inklusive 'Kyssen'?", 'fr': "Quel artiste autrichien a cree des peintures decoratives en feuille d'or dont 'Le Baiser' ?", 'da': "Hvilken østrigsk kunstner skabte dekorative bladguldmalerier, herunder 'Kysset'?"},
 'opt1': {'en': 'Gustav Klimt', 'de': 'Gustav Klimt', 'es': 'Gustav Klimt', 'nl': 'Gustav Klimt', 'sv': 'Gustav Klimt', 'fr': 'Gustav Klimt', 'da': 'Gustav Klimt'},
 'opt2': {'en': 'Egon Schiele', 'de': 'Egon Schiele', 'es': 'Egon Schiele', 'nl': 'Egon Schiele', 'sv': 'Egon Schiele', 'fr': 'Egon Schiele', 'da': 'Egon Schiele'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Caucasus country is home to the ancient city of Tbilisi?', 'de': 'In welchem kaukasischen Land liegt die antike Stadt Tiflis?', 'es': '¿Que pais caucasico alberga la antigua ciudad de Tiflis?', 'nl': 'Welk Kaukasisch land herbergt de oude stad Tbilisi?', 'sv': 'Vilket kaukasiskt land huserar den antika staden Tbilisi?', 'fr': "Quel pays du Caucase abrite l'ancienne cite de Tbilissi ?", 'da': 'Hvilket kaukasisk land huser den antikke by Tbilisi?'},
 'opt1': {'en': 'Georgia', 'de': 'Georgien', 'es': 'Georgia', 'nl': 'Georgie', 'sv': 'Georgien', 'fr': 'la Georgie', 'da': 'Georgien'},
 'opt2': {'en': 'Armenia', 'de': 'Armenien', 'es': 'Armenia', 'nl': 'Armenie', 'sv': 'Armenien', 'fr': "l'Armenie", 'da': 'Armenien'}},

{'category': 'Music', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Finnish composer wrote the tone poem 'Finlandia'?", 'de': "Welcher finnische Komponist schrieb die Tondichtung 'Finlandia'?", 'es': "¿Que compositor finlandes escribio el poema sinfonico 'Finlandia'?", 'nl': "Welke Finse componist schreef het symfonisch gedicht 'Finlandia'?", 'sv': "Vilken finsk kompositor skrev tondiktet 'Finlandia'?", 'fr': "Quel compositeur finlandais a ecrit le poeme symphonique 'Finlandia' ?", 'da': "Hvilken finsk komponist skrev tonedigtet 'Finlandia'?"},
 'opt1': {'en': 'Jean Sibelius', 'de': 'Jean Sibelius', 'es': 'Jean Sibelius', 'nl': 'Jean Sibelius', 'sv': 'Jean Sibelius', 'fr': 'Jean Sibelius', 'da': 'Jean Sibelius'},
 'opt2': {'en': 'Aulis Sallinen', 'de': 'Aulis Sallinen', 'es': 'Aulis Sallinen', 'nl': 'Aulis Sallinen', 'sv': 'Aulis Sallinen', 'fr': 'Aulis Sallinen', 'da': 'Aulis Sallinen'}},

{'category': 'Food', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Iranian dish is saffron-infused rice, often layered with meat or vegetables?', 'de': 'Welches iranische Gericht ist Safran-Reis, oft geschichtet mit Fleisch oder Gemuese?', 'es': '¿Que plato irani es arroz infusionado con azafran, a menudo en capas con carne o verduras?', 'nl': 'Welk Iraans gerecht is met saffraan doortrokken rijst, vaak gelaagd met vlees of groenten?', 'sv': 'Vilken iransk ratt ar saffransriskryddat, ofta lagerlagd med kott eller gronsaker?', 'fr': "Quel plat iranien est du riz au safran, souvent superpose de viande ou de legumes ?", 'da': 'Hvilken iransk ret er safranris, ofte lagdelt med kød eller grøntsager?'},
 'opt1': {'en': 'Tahdig', 'de': 'Tahdig', 'es': 'Tahdig', 'nl': 'Tahdig', 'sv': 'Tahdig', 'fr': 'Tahdig', 'da': 'Tahdig'},
 'opt2': {'en': 'Polo', 'de': 'Polo (Reis)', 'es': 'Polo (arroz)', 'nl': 'Polo (rijst)', 'sv': 'Polo (ris)', 'fr': 'Polo (riz)', 'da': 'Polo (ris)'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Norse mythology, which goddess presides over the afterlife realm named after her?', 'de': 'In der nordischen Mythologie, welche Goettin herrscht ueber das nach ihr benannte Totenreich?', 'es': '¿En la mitologia nordica, que diosa preside el reino del mas alla nombrado en su honor?', 'nl': 'In de Noorse mythologie, welke godin heerst over het naar haar vernoemde dodenrijk?', 'sv': 'I nordisk mytologi, vilken gudinna harskar over dodsriket som ar uppkallat efter henne?', 'fr': "Dans la mythologie nordique, quelle deesse preside le royaume des morts qui porte son nom ?", 'da': 'I nordisk mytologi, hvilken gudinde hersker over dødsriget opkaldt efter hende?'},
 'opt1': {'en': 'Hel', 'de': 'Hel', 'es': 'Hel', 'nl': 'Hel', 'sv': 'Hel', 'fr': 'Hel', 'da': 'Hel'},
 'opt2': {'en': 'Freya', 'de': 'Freya', 'es': 'Freya', 'nl': 'Freya', 'sv': 'Freja', 'fr': 'Freya', 'da': 'Freja'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which actress played a British queen in the Netflix series 'The Crown' during its first two seasons?", 'de': "Welche Schauspielerin spielte eine britische Koenigin in der Netflix-Serie 'The Crown' in den ersten beiden Staffeln?", 'es': "¿Que actriz interpreto a una reina britanica en la serie de Netflix 'The Crown' durante sus dos primeras temporadas?", 'nl': "Welke actrice speelde een Britse koningin in de Netflix-serie 'The Crown' tijdens de eerste twee seizoenen?", 'sv': "Vilken skadespelerska spelade en brittisk drottning i Netflix-serien 'The Crown' under de forsta tva sasongerna?", 'fr': "Quelle actrice a joue une reine britannique dans la serie Netflix 'The Crown' pendant ses deux premieres saisons ?", 'da': "Hvilken skuespiller spillede en britisk dronning i Netflix-serien 'The Crown' i de første to sæsoner?"},
 'opt1': {'en': 'Claire Foy', 'de': 'Claire Foy', 'es': 'Claire Foy', 'nl': 'Claire Foy', 'sv': 'Claire Foy', 'fr': 'Claire Foy', 'da': 'Claire Foy'},
 'opt2': {'en': 'Olivia Colman', 'de': 'Olivia Colman', 'es': 'Olivia Colman', 'nl': 'Olivia Colman', 'sv': 'Olivia Colman', 'fr': 'Olivia Colman', 'da': 'Olivia Colman'}},

{'category': 'Space', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which planet has the shortest year in our solar system, orbiting the sun in just 88 Earth days?', 'de': 'Welcher Planet hat das kuerzeste Jahr in unserem Sonnensystem und umkreist die Sonne in nur 88 Erdentagen?', 'es': '¿Que planeta tiene el año mas corto de nuestro sistema solar, orbitando el sol en solo 88 dias terrestres?', 'nl': 'Welke planeet heeft het kortste jaar in ons zonnestelsel, met een omloop om de zon van slechts 88 aardse dagen?', 'sv': 'Vilken planet har det kortaste aret i vart solsystem, kretsar runt solen pa bara 88 jorddagar?', 'fr': "Quelle planete a l'annee la plus courte de notre systeme solaire, orbitant le soleil en seulement 88 jours terrestres ?", 'da': 'Hvilken planet har det korteste år i vores solsystem, kredser om solen på kun 88 jorddage?'},
 'opt1': {'en': 'Mercury', 'de': 'Merkur', 'es': 'Mercurio', 'nl': 'Mercurius', 'sv': 'Merkurius', 'fr': 'Mercure', 'da': 'Merkur'},
 'opt2': {'en': 'Venus', 'de': 'Venus', 'es': 'Venus', 'nl': 'Venus', 'sv': 'Venus', 'fr': 'Venus', 'da': 'Venus'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which South Caucasus country is home to Mount Ararat's neighboring region?", 'de': 'Welches suedkaukasische Land liegt in der Naehe des Berges Ararat?', 'es': '¿Que pais del Caucaso Sur alberga la region vecina del Monte Ararat?', 'nl': 'Welk Zuid-Kaukasisch land ligt in de buurt van de berg Ararat?', 'sv': 'Vilket sydkaukasiskt land ligger i narheten av berget Ararat?', 'fr': "Quel pays du Caucase du Sud abrite la region voisine du mont Ararat ?", 'da': 'Hvilket sydkaukasisk land huser regionen ved siden af Ararat-bjerget?'},
 'opt1': {'en': 'Armenia', 'de': 'Armenien', 'es': 'Armenia', 'nl': 'Armenie', 'sv': 'Armenien', 'fr': "l'Armenie", 'da': 'Armenien'},
 'opt2': {'en': 'Azerbaijan', 'de': 'Aserbaidschan', 'es': 'Azerbaiyan', 'nl': 'Azerbeidzjan', 'sv': 'Azerbajdzjan', 'fr': "l'Azerbaidjan", 'da': 'Aserbajdsjan'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Japanese designer is known for avant-garde, deconstructed fashion under a black-and-white aesthetic?', 'de': 'Welcher japanische Designer ist fuer avantgardistische, dekonstruierte Mode in schwarz-weisser Ästhetik bekannt?', 'es': '¿Que diseñador japones es conocido por su moda vanguardista y deconstruida bajo una estetica blanco y negro?', 'nl': 'Welke Japanse ontwerper staat bekend om avant-garde, gedeconstrueerde mode in een zwart-witte esthetiek?', 'sv': 'Vilken japansk designer ar kand for avantgardistisk, dekonstruerad mode i svartvit estetik?', 'fr': "Quel createur japonais est connu pour sa mode avant-gardiste et deconstruite dans une esthetique noir et blanc ?", 'da': 'Hvilken japansk designer er kendt for avantgarde, dekonstrueret mode i en sort-hvid æstetik?'},
 'opt1': {'en': 'Yohji Yamamoto', 'de': 'Yohji Yamamoto', 'es': 'Yohji Yamamoto', 'nl': 'Yohji Yamamoto', 'sv': 'Yohji Yamamoto', 'fr': 'Yohji Yamamoto', 'da': 'Yohji Yamamoto'},
 'opt2': {'en': 'Issey Miyake', 'de': 'Issey Miyake', 'es': 'Issey Miyake', 'nl': 'Issey Miyake', 'sv': 'Issey Miyake', 'fr': 'Issey Miyake', 'da': 'Issey Miyake'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which British philosopher wrote 'A Treatise of Human Nature'?", 'de': "Welcher britische Philosoph schrieb 'Ein Traktat ueber die menschliche Natur'?", 'es': "¿Que filosofo britanico escribio 'Tratado de la naturaleza humana'?", 'nl': "Welke Britse filosoof schreef 'Een verhandeling over de menselijke natuur'?", 'sv': "Vilken brittisk filosof skrev 'En avhandling om manskliga naturen'?", 'fr': "Quel philosophe britannique a ecrit 'Traite de la nature humaine' ?", 'da': "Hvilken britisk filosof skrev 'En afhandling om den menneskelige natur'?"},
 'opt1': {'en': 'David Hume', 'de': 'David Hume', 'es': 'David Hume', 'nl': 'David Hume', 'sv': 'David Hume', 'fr': 'David Hume', 'da': 'David Hume'},
 'opt2': {'en': 'John Locke', 'de': 'John Locke', 'es': 'John Locke', 'nl': 'John Locke', 'sv': 'John Locke', 'fr': 'John Locke', 'da': 'John Locke'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which quilting pattern uses a repeated square-in-square design, often called a nine-patch variant?', 'de': 'Welches Quiltmuster verwendet ein wiederholtes Quadrat-im-Quadrat-Design, oft als Nine-Patch-Variante bezeichnet?', 'es': '¿Que patron de acolchado usa un diseño repetido de cuadrado dentro de cuadrado, a menudo llamado variante de nueve parches?', 'nl': 'Welk quiltpatroon gebruikt een herhaald vierkant-in-vierkant ontwerp, vaak een nine-patch-variant genoemd?', 'sv': 'Vilket lapptacksmonster anvander en upprepad kvadrat-i-kvadrat-design, ofta kallad en nio-lapps-variant?', 'fr': "Quel motif de patchwork utilise un design repete de carre dans un carre, souvent appele variante du 'nine-patch' ?", 'da': 'Hvilket quiltmønster bruger et gentaget kvadrat-i-kvadrat-design, ofte kaldet en ni-lap-variant?'},
 'opt1': {'en': 'Nine Patch', 'de': 'Nine Patch', 'es': 'Nueve parches', 'nl': 'Nine Patch', 'sv': 'Nio-lapps', 'fr': 'Nine Patch', 'da': 'Ni-lap'},
 'opt2': {'en': 'Log Cabin', 'de': 'Blockhuette', 'es': 'Cabaña de troncos', 'nl': 'Blokhut', 'sv': 'Timmerstuga', 'fr': 'Cabane en rondins', 'da': 'Log Cabin'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which flowering plant's fragrant blooms are used to make traditional rosewater?", 'de': 'Welche bluehende Pflanze liefert die duftenden Blueten fuer traditionelles Rosenwasser?', 'es': '¿Que planta con flores tiene flores fragantes usadas para hacer el agua de rosas tradicional?', 'nl': 'Welke bloeiende plant heeft geurige bloesems die worden gebruikt om traditioneel rozenwater te maken?', 'sv': 'Vilken blommande vaxts doftande blommor anvands for att gora traditionellt rosvatten?', 'fr': "Quelle plante a fleurs a des fleurs parfumees utilisees pour faire l'eau de rose traditionnelle ?", 'da': 'Hvilken blomstrende plantes duftende blomster bruges til at lave traditionelt rosenvand?'},
 'opt1': {'en': 'Damask rose', 'de': 'Damaszenerrose', 'es': 'Rosa de Damasco', 'nl': 'Damascusroos', 'sv': 'Damascenerros', 'fr': 'Rose de Damas', 'da': 'Damascener-rose'},
 'opt2': {'en': 'Jasmine', 'de': 'Jasmin', 'es': 'Jazmin', 'nl': 'Jasmijn', 'sv': 'Jasmin', 'fr': 'Jasmin', 'da': 'Jasmin'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which architect designed the Farnsworth House, a glass-walled modernist residence?', 'de': 'Welcher Architekt entwarf das Farnsworth House, ein modernistisches Wohnhaus mit Glaswaenden?', 'es': '¿Que arquitecto diseño la Casa Farnsworth, una residencia modernista de paredes de vidrio?', 'nl': 'Welke architect ontwierp het Farnsworth House, een modernistische woning met glazen wanden?', 'sv': 'Vilken arkitekt designade Farnsworth House, en modernistisk bostad med glasvaggar?', 'fr': "Quel architecte a conçu la Farnsworth House, une residence moderniste aux murs de verre ?", 'da': 'Hvilken arkitekt designede Farnsworth House, en modernistisk bolig med glasvægge?'},
 'opt1': {'en': 'Ludwig Mies van der Rohe', 'de': 'Ludwig Mies van der Rohe', 'es': 'Ludwig Mies van der Rohe', 'nl': 'Ludwig Mies van der Rohe', 'sv': 'Ludwig Mies van der Rohe', 'fr': 'Ludwig Mies van der Rohe', 'da': 'Ludwig Mies van der Rohe'},
 'opt2': {'en': 'Philip Johnson', 'de': 'Philip Johnson', 'es': 'Philip Johnson', 'nl': 'Philip Johnson', 'sv': 'Philip Johnson', 'fr': 'Philip Johnson', 'da': 'Philip Johnson'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Norse mythology, which world tree connects the nine realms?', 'de': 'In der nordischen Mythologie, welcher Weltenbaum verbindet die neun Reiche?', 'es': '¿En la mitologia nordica, que arbol del mundo conecta los nueve reinos?', 'nl': 'In de Noorse mythologie, welke wereldboom verbindt de negen rijken?', 'sv': 'I nordisk mytologi, vilket varldstrad forbinder de nio rikena?', 'fr': "Dans la mythologie nordique, quel arbre-monde relie les neuf royaumes ?", 'da': 'I nordisk mytologi, hvilket verdenstræ forbinder de ni riger?'},
 'opt1': {'en': 'Yggdrasil', 'de': 'Yggdrasil', 'es': 'Yggdrasil', 'nl': 'Yggdrasil', 'sv': 'Yggdrasil', 'fr': 'Yggdrasil', 'da': 'Yggdrasil'},
 'opt2': {'en': 'Bifrost', 'de': 'Bifroest', 'es': 'Bifrost', 'nl': 'Bifrost', 'sv': 'Bifrost', 'fr': 'Bifrost', 'da': 'Bifrost'}},
]
