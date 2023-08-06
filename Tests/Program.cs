


string st = "Casting Time 1 standard action\r\n" + " Components V,S";

string[] castArray = st.Split('\n');

string castingTime = castArray[0].Substring(14);

string components = castArray[1].Substring(11);

Console.WriteLine("Casting time" + castingTime + "   "  +  "componentes" + components);