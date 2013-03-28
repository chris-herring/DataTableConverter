DataTableConverter
==================

A Newtonsoft Json converter to the .NET System.Data.DataTable

Usage
=====

Include the convert into your project.

Then when serializing/deserializing an object include the converter

<code>
string json = JsonConvert.SerializeObject(extract.Data, new Serialization.DataTableConverter());
DataTable table = JsonConvert.DeserializeObject<DataTable>(json, new Serialization.DataTableConverter());
</code>

<code>
Or you can edit the global configuration for the JSON .NET serializer so that the converter will be picked up everywhere:

GlobalConfiguration.Configuration.Formatters.JsonFormatter.SerializerSettings.Converters.Add(new Serialization.DataTableConverter());
</code>