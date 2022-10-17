using System.Runtime.InteropServices;

[ComVisible(true)]
public class TestClass
{

    public TestClass(string name, int age)
    {
        Name = name;
        Age = age;
    }

    public string Name { get; set; }
    public int Age { get; set; }

    public override string ToString() => $"{Name} {Age}";
}
