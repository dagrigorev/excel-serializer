namespace DemoApp;

public class SampleClass
{
    public string Name { get; set; }
    public int Age { get; set; }
    public decimal Balance { get; set; }
    public string? Description { get; set; }

    public override string ToString()
    {
        return $"Name: {Name}, Age: {Age}, Balance: {Balance}, Description: {Description ?? "N/A"}";
    }
}
