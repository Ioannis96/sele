using OpenQA.Selenium.DevTools.V112.Debugger;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sele
{

    public class Spell
    {
        public string Name { get; set; }
        public string? School { get; set; }
        public List<string>? Levels { get; set; }
        public List<string>? Components { get; set; }
        public string? CastingTime { get; set; }
        public string? Range { get; set; }
        public string? Target { get; set; }
        public string? Effect { get; set; }
        public string? Area { get; set; }
        public string? Duration { get; set; }
        public string? SavingThrow { get; set; }
        public string? SpellResistance { get; set; }
        public string? Description { get; set; }

        public Spell(string textName)
        {
            Name = textName;
        }



    }
}
