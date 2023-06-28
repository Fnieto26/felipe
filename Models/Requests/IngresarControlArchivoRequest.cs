namespace Directions_Api.Models.Requests
{
    public class IngresarControlArchivoRequest
    {
        public string Usuario { get; set; }
        public string Contrasena { get; set; }
        public string NombreArchivo { get; set; }
        public long TamanoArchivo { get; set; }
        public int Aproximacion { get; set; }
    }
}
