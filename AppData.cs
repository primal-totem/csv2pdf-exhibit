using System;

public class AppData
{
    public string name { get; set; }
    public string description { get; set; }
    public string photo { get; set; }
    public AppData(string name, string description, string photo)
    {
        this.name = name;
        this.description = description;
        this.photo = photo;
    }
}
