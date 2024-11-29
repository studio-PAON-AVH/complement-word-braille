﻿
using fr.avh.braille.dictionnaire;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace fr.avh.braille.tests
{
    [TestClass]
    public class TestAbreviation
    {
        // TODO : reprendre les listes de mots en erreur de détection
        // et faire des tests unitaires pour les mots en erreur
        [TestMethod]
        [DataRow("chamallow", true)]
        [DataRow("about", true)]
        [DataRow("Armanet", true)]
        [DataRow("between", true)]
        [DataRow("camelli", true)]
        [DataRow("cliffhanger", true)]
        [DataRow("Colagreco", true)]
        [DataRow("crowd", true)]
        [DataRow("Danemark", true)]
        [DataRow("Decitre", true)]
        [DataRow("Desplat", true)]
        [DataRow("directors", true)]
        [DataRow("Droopy", true)]
        [DataRow("Electric", true)]
        [DataRow("española", true)]        
        [DataRow("Fabrice", true)]
        [DataRow("Fenoglio", true)]
        [DataRow("found", true)]
        [DataRow("friend", true)]
        [DataRow("Garrett", true)]
        [DataRow("Ghibli", true)]
        [DataRow("Goldin", true)]
        [DataRow("Habré", true)]
        [DataRow("Harouth", true)]
        [DataRow("Immagine", true)]
        [DataRow("insider", true)]
        [DataRow("Jaclyn", true)]
        [DataRow("Jee-woon", true)]
        [DataRow("Jurassic", true)]
        [DataRow("Kalapapruek", true)]
        [DataRow("looking", true)]
        [DataRow("Lubrick", true)]
        [DataRow("Macri", true)]
        [DataRow("Magre", true)]
        [DataRow("Magris", true)]
        [DataRow("minorities", true)]
        [DataRow("Nebraska", true)]
        [DataRow("never", true)]
        [DataRow("Nostromo", true)]
        [DataRow("Opening", true)]
        [DataRow("Pablo", true)]
        [DataRow("Padre", true)]
        [DataRow("Padrone", true)]
        [DataRow("Patry", true)]
        [DataRow("Petronio", true)]
        [DataRow("Petrovic", true)]
        [DataRow("Pierre-Yves", true)]
        [DataRow("retakes", true)]
        [DataRow("revivals", true)]
        [DataRow("Ritrovata", true)]
        [DataRow("Ritrovato", true)]
        [DataRow("Robledo", true)]
        [DataRow("Siegfried", true)]
        [DataRow("Sinicropi", true)]
        [DataRow("Szafran", true)]
        [DataRow("Szifrón", true)]
        [DataRow("Tetro", true)]
        [DataRow("UCLA", true)]
        [DataRow("Vignoli", true)]
        [DataRow("watch", true)]
        [DataRow("wishes", true)]
        [DataRow("wonderboy", true)]
        [DataRow("wrong", true)]
        [DataRow("you", true)]
        [DataRow("Henri", true)]
        [DataRow("Colombie", true)]
        public void MotsEsAbregeable(string mot, bool expected)
        {
            Console.WriteLine("Détection et syllabes : {0} ", Abreviation.regleAppliquerSur(mot));
            Assert.AreEqual(
                expected,
                Abreviation.EstAbregeable(mot),
                "Mot:<{0}>",
                new object[] { mot });
        }

        [TestMethod]
        [DataRow("Afghanistan", false)]
        [DataRow("Amnesty", false)]
        [DataRow("Biarritz", false)]
        [DataRow("Bucarest", false)]
        [DataRow("Budapest", false)]
        [DataRow("Cinespace", false)]
        [DataRow("Comodoro", false)]
        [DataRow("Comoedia", false)]
        [DataRow("Everest", false)]
        [DataRow("Forest", false)]
        [DataRow("Foresti", false)]
        [DataRow("Galeshka", false)]
        [DataRow("Giono", false)]
        [DataRow("Happiest", false)]
        [DataRow("Jean-Luc", false)]
        [DataRow("Jean-Marie", false)]
        [DataRow("Kassovitz", false)]
        [DataRow("Kaufman", false)]
        [DataRow("Kracauer", false)]
        [DataRow("Lazarescu", false)]
        [DataRow("Limonest", false)]
        [DataRow("Lionel", false)]
        [DataRow("Maïwenn", false)]
        [DataRow("Majestic", false)]
        [DataRow("man's", false)]
        [DataRow("Oakland", false)]
        [DataRow("Olaciragui", false)]
        [DataRow("Palestine", false)]
        [DataRow("Stazione", false)]
        [DataRow("successful", false)]
        [DataRow("Witness", false)]
        public void MotsEsNonAbregeable(string mot, bool expected)
        {
            Console.WriteLine("Détection et syllabes : {0} ", Abreviation.regleAppliquerSur(mot));
            Assert.AreEqual(
                expected,
                Abreviation.EstAbregeable(mot),
                "Mot:<{0}>",
                new object[] { mot });
        }
    }
}
