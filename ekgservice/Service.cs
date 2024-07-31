using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceProcess;
using System.Timers;
using System.IO;
using System.Data.SqlClient;
namespace ekgservice
{
    class Service : ServiceBase
    {
        private Timer timer;

        public Service()
        {
            this.ServiceName = "Ekgservice";
            this.EventLog.Source = "EkgserviceSouce";
            this.EventLog.Log = "EkgserviceLog";
            /*if (!EventLog.SourceExists("EkgserviceSouce"))
            {
                EventLog.CreateEventSource("EkgserviceSouce", "EkgserviceLog");
            }*/
        }
        protected override void OnStart(string[] args)
        {
            //base.OnStart(args);
            EventLog.WriteEntry("EkgService a iniciado");
            timer = new Timer(300000);
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
        }

        protected override void OnStop()
        {
            EventLog.WriteEntry("EkgService se ha detenido.");
            timer.Stop();
        }

        private void OnTimer(object sender, ElapsedEventArgs e)
        {
            string ruta = @"";
            try
            {
                if (Directory.Exists(ruta))
                {
                    ExplorarDirectorio(ruta);
                }
                else
                {
                    EventLog.WriteEntry("La ruta especificada no existe.");
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("Se produjo un error: " + ex.Message);
            }
        }

        public static void Main()
        {
            ServiceBase.Run(new Service());
        }
        private void ExplorarDirectorio(string ruta)
        {
            try
            { 
                string[] archivos = Directory.GetFiles(ruta);
                string[] subdirectorios = Directory.GetDirectories(ruta);
                string connectionString = "";
                if (archivos.Length > 0)
                {
                    foreach (var archivo in archivos)
                    {
                        string rutaOrigen = Path.Combine(ruta, Path.GetFileName(archivo));
                        string nombre = Path.GetFileName(archivo);
                        string[] nombre_ = nombre.Split('.');
                        string[] subcadenas = nombre_[0].Split('_');
                        string doc = subcadenas[subcadenas.Length-1];
                        //CONSULTA 
                        string query = "";
                        if (!string.IsNullOrEmpty(doc))
                        {
                            using (SqlConnection connection = new SqlConnection(connectionString))
                            {
                                try
                                {
                                    connection.Open();
                                    SqlCommand command = new SqlCommand(query, connection);
                                    command.Parameters.AddWithValue("@doc", doc);

                                    using (SqlDataReader reader = command.ExecuteReader())
                                    {
                                        if (reader.Read())
                                        {
                                            int IdProcedimiento = reader.GetInt32(0);
                                            int idinforme;
                                            //CONSULTA DEL ULTIMO CORRELATIVO EN SS_PR_PROCEDIMIENTOINFORME
                                            string query2 = "C";
                                            using (SqlConnection connection2 = new SqlConnection(connectionString))
                                            {
                                                try
                                                {
                                                    connection2.Open();
                                                    SqlCommand command2 = new SqlCommand(query2, connection2);
                                                    command2.Parameters.AddWithValue("@idproc", IdProcedimiento);
                                                    using (SqlDataReader reader2 = command2.ExecuteReader())
                                                    {
                                                        idinforme = reader2.Read() ? reader2.GetInt32(0) : 0;
                                                        //MOVER LOS ARCHIVOS
                                                        string newDestino = @"" + nombre;
                                                        try
                                                        {
                                                            File.Move(rutaOrigen, newDestino);

                                                            string query3 = "";
                                                            using (SqlConnection connection3 = new SqlConnection(connectionString))
                                                            {
                                                                SqlCommand command3 = new SqlCommand(query3, connection3);
                                                                DateTime now = DateTime.Now;
                                                                command3.Parameters.AddWithValue("@IdProcedimiento", IdProcedimiento);
                                                                command3.Parameters.AddWithValue("@IdInforme", idinforme + 4);
                                                                command3.Parameters.AddWithValue("@Nombre", "EKG E");
                                                                command3.Parameters.AddWithValue("@RutaInforme", newDestino);
                                                                command3.Parameters.AddWithValue("@Estado", 2);
                                                                command3.Parameters.AddWithValue("@UsuarioCreacion", "AUTOMATICO");
                                                                command3.Parameters.AddWithValue("@FechaCreacion", now);
                                                                command3.Parameters.AddWithValue("@UsuarioModificacion", "AUTOMATICO");
                                                                command3.Parameters.AddWithValue("@FechaModificacion", now);
                                                                connection3.Open();
                                                                int result = command3.ExecuteNonQuery();
                                                                EventLog.WriteEntry("Cantidad de resultados: " + result.ToString() + " insertados para el procedimiento "+IdProcedimiento+"\r\n");
                                                            }
                                                        }
                                                        catch (IOException ioEx)
                                                        {
                                                            //LOG
                                                            EventLog.WriteEntry(ioEx.Message + "\r\n");
                                                        }
                                                        catch (UnauthorizedAccessException unAuthEx)
                                                        {
                                                            // Maneja los errores de acceso no autorizado
                                                            EventLog.WriteEntry(unAuthEx.Message + "\r\n");
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            // Maneja cualquier otro tipo de error
                                                            EventLog.WriteEntry(ex.Message + "\r\n");
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    EventLog.WriteEntry(ex.Message + "\r\n");
                                                }
                                            }
                                        }
                                        else
                                        {
                                            EventLog.WriteEntry("Procedimiento asociado al documento " + doc + " no encontrado en la base de datos \r\n");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    EventLog.WriteEntry(ex.Message+"\r\n");
                                }
                            }
                        }
                        else
                        {
                            EventLog.WriteEntry("Documento null o vacio \r\n");
                        }
                    }
                }
                else
                {
                        EventLog.WriteEntry("No hay archivos en "+ruta+"\r\n");
                }

                foreach (var subdirectorio in subdirectorios)
                {
                    ExplorarDirectorio(subdirectorio); // Llamada recursiva para explorar subdirectorios
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("Error al acceder a {ruta}: {ex.Message}\r\n");
            }
        }
    }
}
