using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Microsoft.Data.Sql;
using System.Windows.Forms;

namespace Printickets
{

    public partial class Form1 : Form
    {

        //Variables para la Base de datos
        Decimal canti;
        String Codex;
        String d1;
        String d2;
        String origen;
        String factura;
        String lane;
        String Pin;
        Decimal OrderM;
        

        //Variable para la cantidad de etiquetas a imprimir
        int cEtiquetas;

        //Variables para el zpl y send to printer
        String[] zpl;
        String toPrint;
        String toSend;

        //Etiqeuta para la impresora 
                //END terminacion de etiqueta
        String end = "^XZ";
                
        String start = "^XA\r\n"
                     + "^CI28^FS\r\n";
        String top = "^FO655,20^GB10,350,10^FS\r\n"
                     + "^FO660,20^AJR,38^FD Distribution Center^FS\r\n";
        String cubo = "^FO235,850^CF0,60\r\n"
                     + "^GB50,170,50,,^FS\r\n"
                     + "^FO150,850^GB130,170,10^FS\r\n"
                     + "^FO240,860^ADR,16,10^AJ,36\r\n"
                     + "^FR^FD ORIGEN ^FS\r\n";

        //Variables para obtener datos de la tabla y agregar a zpl

        String ca;
        String Codigo;
        String descr1;
        String Descr2;
        String lna;
        String or;
        String p;
        String Ord;
        String f;
        

        public Form1()
        {
            InitializeComponent();
        }

       
        
        //Funcion que manda a imprimir
        public void sendToPrint(String stp)
        {
            String ip = "192.10.10.11";
            int port = 9100;

            try
            {
                System.Net.Sockets.TcpClient client = new System.Net.Sockets.TcpClient();

                client.Connect(ip, port);
                System.IO.StreamWriter writter = new System.IO.StreamWriter(client.GetStream());
                 writter.Write(stp);
              
                writter.Flush();


                client.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Fallo conexion a Impresora: " + ex.Message);
            }
        }

        //Funcion para generar etiquta
        public String generarZPL(string [] zpl)
        {
            toPrint = "";
            toPrint += start + top + cubo;

            //Codigo
            if (zpl[1].Length >= 15)
            {
               
                switch (zpl[1].Length)
                {
                    case 15:  
                        toPrint += "^FO500,10^APR,170,170^FD" + zpl[1] + "^FS \r\n"; 
                        break;

                    case 16:
                        
                        toPrint += "^FO500,10^APR,170,130^FD" + zpl[1] + "^FS \r\n";
                        break;
                    case 17:
                        
                        toPrint += "^FO500,10^APR,170,130^FD" + zpl[1] + "^FS \r\n";
                        break;
                    case 18:
                        
                        toPrint += "^FO500,10^APR,170,110^FD" + zpl[1] + "^FS \r\n";
                        break;
                    case 19:
                        
                        toPrint += "^FO500,10^APR,170,110^FD" + zpl[1] + "^FS \r\n";
                        break;
                    case 20:
                       
                        toPrint += "^FO500,10^APR,170,130^FD" + zpl[1] + "^FS \r\n";
                        break;
                    case 21:
                        
                        toPrint += "^FO500,10^APR,170,150^FD" + zpl[1] + "^FS \r\n";
                        break;
                }

                

            }
            else
            {
                
                toPrint += "^FO500,10^APR,170,190^FD" + zpl[1] + "^FS \r\n";
            }
            

           

            

            //Largo del cubo negro
            switch (zpl[2].Length)
            {
                case 5:
                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,200,100,,^FS\r\n";
                    break;
                case 6:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,200,100,,^FS\r\n";
                    break;
                case 7:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,215,100,,^FS\r\n";
                    break;
                case 8:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,225,100,,^FS\r\n";
                    break;
                case 9:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,230,100,,^FS\r\n";
                    break;
                case 10:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,280,90,,^FS\r\n";
                    break;
                case 11:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,300,90,,^FS\r\n";
                    break;
                case 12:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,330,90,,^FS\r\n";
                    break;
                case 13:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,360,90,,^FS\r\n";
                    break;
                case 14:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,390,90,,^FS\r\n";
                    break;
                case 15:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,410,90,,^FS\r\n";
                    break;
                case 16:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,430,90,,^FS\r\n";
                    break;
                case 17:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,450,90,,^FS\r\n";
                    break;
                case 18:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,470,90,,^FS\r\n";
                    break;
                case 19:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,500,90,,^FS\r\n";
                    break;
                case 20:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,540,90,,^FS\r\n";
                    break;
                case 21:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,560,90,,^FS\r\n";
                    break;
                case 22:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,585,90,,^FS\r\n";
                    break;
                case 23:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,620,90,,^FS\r\n";
                    break;

                case 24:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,645,90,,^FS\r\n";
                    break;
                case 25:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,680,90,,^FS\r\n";
                    break;
                case 26:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,700,90,,^FS\r\n";
                    break;
                case 27:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,750,90,,^FS\r\n";
                    break;
                case 28:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,770,90,,^FS\r\n";
                    break;
                case 29:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,780,90,,^FS\r\n";
                    break;
                case 30:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,790,90,,^FS\r\n";
                    break;
                case 31:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,800,90,,^FS\r\n";
                    break;
                case 32:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,810,90,,^FS\r\n";
                    break;
                case 33:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB115,840,90,,^FS\r\n";
                    break;
                case 34:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB105,940,90,,^FS\r\n";
                    break;
                case 35:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB105,960,90,,^FS\r\n";
                    break;
                case 36:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB105,980,90,,^FS\r\n";
                    break;
                case 37:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB105,990,90,,^FS\r\n";
                    break;
                case 38:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB105,1000,90,,^FS\r\n";
                    break;
                case 39:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB105,1040,90,,^FS\r\n";
                    break;
                case 40:

                    toPrint += "^FO415,20^CF0,70,\r\n"
                              + "^GB105,980,90,,^FS\r\n";
                    break;
            }
            //Description 1
            toPrint += "^FO395,20^APR,120,60\r\n"
                    + "^FR^FD" + zpl[2] + "^FS\r\n";

            //Description 2
            toPrint += "^FO265,20^ADR,16,10^AJ,120,60^FD" + zpl[3] + "^FS\r\n";

            //Codigo de barras
            if (Convert.ToString(zpl[1]) == "@N-2")
            {
                toPrint += "^FO160,870^ADR,16,10^AJ,70,35\r\n"
                           + "^FR^FD" + zpl[5] + "^FS\r\n";
                if (zpl[1].Length >= 17)
                {
                    toPrint += "^FX CODIGO DE BARRAS\r\n"
                           + "^FO180,25^BY3,3,10\r\n"
                           + "^BCR,100,N,N\r\n"
                           + "^FD" + zpl[1] + "@N/2^FS\r\n";
                }
                else
                {
                     toPrint += "^FX CODIGO DE BARRAS\r\n"
                           + "^FO180,25^BY4\r\n"
                           + "^BCR,100,N,N\r\n"
                           + "^FD" + zpl[1] + "@N/2^FS\r\n";
                }
                   
            }
            else if (Convert.ToString(zpl[5]) == "2")
            {
                toPrint += "^FO160,910^ADR,16,10^AJ,70,35\r\n"
                           + "^FR^FD" + zpl[5] + "^FS\r\n";
                if (zpl[1].Length >= 17)
                {
                    toPrint += "^FX CODIGO DE BARRAS\r\n"
                           + "^FO180,25^BY3,2,10\r\n"
                           + "^BCR,100,N,N\r\n"
                           + "^FD" + zpl[1] + "-2^FS\r\n";
                }
                else
                {
                    toPrint += "^FX CODIGO DE BARRAS\r\n"
                           + "^FO180,25^BY3,2,10\r\n"
                           + "^BCR,100,N,N\r\n"
                           + "^FD" + zpl[1] + "-2^FS\r\n";
                }
                    
                
            }
            else if (Convert.ToString(zpl[5]) == "1")
            {
                toPrint += "^FO160,910^ADR,16,10^AJ,70,35\r\n"
                          + "^FR^FD" + zpl[5] + "^FS\r\n";

                if (zpl[1].Length >= 17)
                {
                    toPrint += "^FX CODIGO DE BARRAS\r\n"
                           + "^FO180,25^BY3,2,10\r\n"
                           + "^BCR,100,N,N\r\n"
                           + "^FD" + zpl[1] + "-1^FS\r\n";
                }
                else
                {
                    toPrint += "^FX CODIGO DE BARRAS\r\n"
                           + "^FO180,25^BY3,2,10\r\n"
                           + "^BCR,100,N,N\r\n"
                           + "^FD" + zpl[1] + "-1^FS\r\n";
                }
                    
            }
            else
            {
                toPrint += "^FO160,910^ADR,16,10^AJ,70,35\r\n"
                           + "^FR^FD" + zpl[5] + "^FS\r\n";
                if (zpl[1].Length >= 17)
                {
                    toPrint += "^FX CODIGO DE BARRAS\r\n"
                           + "^FO180,25^BY3,2,10\r\n"
                           + "^BCR,100,N,N\r\n"
                           + "^FD" + zpl[1] + "^FS\r\n";
                }
                else
                {
                    toPrint += "^FX CODIGO DE BARRAS\r\n"
                           + "^FO180,25^BY3,2,10\r\n"
                           + "^BCR,100,N,N\r\n"
                           + "^FD" + zpl[1] + "^FS\r\n";
                }
                    
                
            }

            //Linea

           

            //Pin 
            if (zpl[6].Length > 12)
            {
               
                if(zpl[6].Length == 24)
                {
                    toPrint += "^FO125,10^AJR,26,18^FD P/N:" + zpl[6] + " ^FS\r\n";
                    toPrint += "^FO125,340^AJR,26^FD Linea:" + Convert.ToString(zpl[4]) + " ^FS\r\n";

                }
                else
                {
                    toPrint += "^FO125,10^AJR,26,22^FD P/N:" + zpl[6] + " ^FS\r\n";
                    toPrint += "^FO125,250^AJr,26^FD Linea:" + Convert.ToString(zpl[4]) + " ^FS\r\n";
                }
            }
            else
            {
                toPrint += "^FO125,10^AJR,26^FD P/N:" + zpl[6] + " ^FS\r\n";
                toPrint += "^FO125,250^AJR,26^FD Linea:" + Convert.ToString(zpl[4]) + " ^FS\r\n";

            }


            // Factura Unit Per Box
            toPrint += "^FO125,570^AJR,26^FD Factura:" + zpl[7] + " ^FS\r\n";
            toPrint += "^FO655,690^AJR,32,32^FD Units Per Box:" + zpl[8] + "^FS\r\n";
            toPrint += end;

            return toPrint;
        }
        //Funcion que no hace nada pero si se borra falla
        private void label1_Click(object sender, EventArgs e)
        {

        }

        // Funcion de boton (Facturas, Entrada merc, pedidos)
        private void button2_Click(object sender, EventArgs e)
        {
            // guarda el valor que se escribe en el capo de texto 
            factura = textBox1.Text;
            //compara si se escribio algo o no
            if (factura != "")
            {
                //hace visible el boton imprimir
                print.Visible = true;
                //hace visible la tabla
                dataGridView1.Visible = true;

                //Limita la cantidad de caracteres en descripciones
                ((DataGridViewTextBoxColumn)dataGridView1.Columns[2]).MaxInputLength = 37;
                ((DataGridViewTextBoxColumn)dataGridView1.Columns[3]).MaxInputLength = 31;

                //limpia los datos de la tabla
                dataGridView1.Rows.Clear();

                try
                {
                    //datos para conectar a base de datos
                    String str = "server=SERVER; database=SBO_TRANSMISIONES;UID=consultor;password=Consult0r*";
                    //Query 
                    String qry = "SELECT [SBO_TRANSMISIONES].[dbo].POR1.Quantity, " +
                        "[SBO_TRANSMISIONES].[dbo].POR1.ItemCode," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.ItemName," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.FrgnName," +
                        "[SBO_TRANSMISIONES].[dbo].OITB.ItmsGrpNam," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.OrdrMulti," +
                        "[SBO_TRANSMISIONES].[dbo].OPOR.NumAtCard," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.SuppCatNum " +
                        "FROM POR1 INNER JOIN OPOR ON POR1.DocEntry = OPOR.DocEntry INNER JOIN OITM ON POR1.ItemCode = OITM.ItemCode INNER JOIN OITB ON OITM.ItmsGrpCod= OITB.ItmsGrpCod WHERE OPOR.NumAtCard ='" + factura + "'";
                    //CON variable que guarda la conexion a BD
                    SqlConnection con = new SqlConnection(str);
                    //CMD variable que guarda el query
                    SqlCommand cmd = new SqlCommand(qry, con);
                    //Abre conexion a la BD
                    con.Open();
                    //Ejecuta el query
                    SqlDataReader reader = cmd.ExecuteReader();

                    //Lee cada fila de la tabla mientras no este vacia, si esta vacia se acaba la tabla
                    while (reader.Read())
                    {
                        //Obtener los valores de cada fila una vez 
                        canti = Convert.ToInt32(reader.GetDecimal(0));
                        Codex = reader.GetString(1);
                        d1 = reader.GetString(2);
                        d2 = reader.GetString(3);
                        lane = reader.GetString(4);
                        OrderM = Convert.ToInt32(reader.GetDecimal(5));
                        if (OrderM == 0)
                            OrderM = 1;
                        factura = reader.GetString(6);

                      

                        //OPBTENER PIN SI SUPPCATNUM IS NULL
                        if (reader.IsDBNull(7))
                        {
                            Pin = Codex;

                        }
                        else
                        {
                            Pin = reader.GetString(7);
                        }

                        //Origen
                        if (Codex.IndexOf("-") != -1)
                        {
                            if (Codex.IndexOf("@") != -1)
                            {
                                origen = Codex.Substring(Codex.IndexOf("@"));
                                Codex = Codex.Substring(0, Codex.IndexOf("@"));
                            }
                            else
                            {
                                origen = Codex.Substring(Codex.IndexOf("-") + 1);
                                Codex = Codex.Substring(0, Codex.IndexOf("-"));
                            }
                        }
                        else
                        {
                            origen = "0";
                        }

                        //LIMITAR DESCRIPCION POR DEFAULT
                        if (d1.Length > 38)
                        {
                            d1 = d1.Substring(0, 37);
                          
                        }

                        if (d2.Length > 31)
                        {
                            d2 = d2.Substring(0, 30);
                           
                        }
                        //ROW variable de la fila de la tabla qeu se va a "pintar"
                        //row.cells[n] coloca un valor en cada espacio de la fila
                        DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                        row.Cells[0].Value = canti;
                        row.Cells[1].Value = Codex;
                        row.Cells[2].Value = d1;
                        row.Cells[3].Value = d2;
                        row.Cells[4].Value = lane;
                        row.Cells[5].Value = origen;
                        row.Cells[7].Value = Pin;
                        row.Cells[8].Value = factura;
                        row.Cells[9].Value = OrderM;



                        //Agregar datos a la tabla
                        dataGridView1.Rows.Add(row);
                    }
                }//Si falla en algun punto este codigo se ejecutara
                catch (Exception es)
                {   //Mensaje del error
                    MessageBox.Show("Error en conexion o creacion de tabla (Consultar con Sistemas)" + es.Message);
                }
            }
        }

        // Funcion de boton (Facturas, Entrada merc, pedidos)
        private void button1_Click(object sender, EventArgs e)
        {
            factura = textBox1.Text;
            if(factura!= "")
            {
                print.Visible = true;
                dataGridView1.Visible = true;

                ((DataGridViewTextBoxColumn)dataGridView1.Columns[2]).MaxInputLength = 37;
                ((DataGridViewTextBoxColumn)dataGridView1.Columns[3]).MaxInputLength = 31;

                dataGridView1.Rows.Clear();

                try
                {
                    String str = "server=SERVER; database=SBO_TRANSMISIONES;UID=consultor;password=Consult0r*";
                    String qry = "SELECT [SBO_TRANSMISIONES].[dbo].PDN1.Quantity, " +
                        "[SBO_TRANSMISIONES].[dbo].PDN1.ItemCode," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.ItemName," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.FrgnName," +
                        "[SBO_TRANSMISIONES].[dbo].OITB.ItmsGrpNam," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.OrdrMulti," +
                        "[SBO_TRANSMISIONES].[dbo].opdn.NumAtCard," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.SuppCatNum " +
                        "FROM PDN1 INNER JOIN OPDN ON PDN1.DocEntry = OPDN.DocEntry INNER JOIN OITM ON PDN1.ItemCode = OITM.ItemCode INNER JOIN OITB ON OITM.ItmsGrpCod= OITB.ItmsGrpCod WHERE OPDN.NumAtCard ='" + factura + "'";
                    SqlConnection con = new SqlConnection(str);
                    SqlCommand cmd = new SqlCommand(qry, con);
                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();


                    while (reader.Read())
                    {
                        canti = Convert.ToInt32(reader.GetDecimal(0));
                        Codex = reader.GetString(1);
                        d1 = reader.GetString(2);
                        d2 = reader.GetString(3);
                        lane = reader.GetString(4);
                        OrderM = Convert.ToInt32(reader.GetDecimal(5));
                        if (OrderM == 0)
                            OrderM = 1;
                        factura = reader.GetString(6);


                      
                        //Origen
                        if (Codex.IndexOf("-") != -1)
                        {
                            if (Codex.IndexOf("@") != -1)
                            {
                                origen = Codex.Substring(Codex.IndexOf("@"));
                                Codex = Codex.Substring(0, Codex.IndexOf("@"));
                            }
                            else
                            {
                                origen = Codex.Substring(Codex.IndexOf("-") + 1);
                                Codex = Codex.Substring(0, Codex.IndexOf("-"));
                            }
                        }
                        else
                        {
                            origen = "0";
                        }
                        //OPBTENER PIN SI SUPPCATNUM IS NULL
                        if (reader.IsDBNull(7))
                        {
                            Pin = Codex;

                        }
                        else
                        {
                            Pin = reader.GetString(7);
                        }
                        //LIMITAR DESCRIPCION POR DEFAULT
                        if (d1.Length > 38)
                        {
                            d1 = d1.Substring(0, 38);
                            
                        }

                        if (d2.Length > 31)
                        {
                            d2 = d2.Substring(0, 31);
                            
                        }
                        DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                        row.Cells[0].Value = canti;
                        row.Cells[1].Value = Codex;
                        row.Cells[2].Value = d1;
                        row.Cells[3].Value = d2;
                        row.Cells[4].Value = lane;
                        row.Cells[5].Value = origen;
                        row.Cells[7].Value = Pin;
                        row.Cells[8].Value = factura;
                        row.Cells[9].Value = OrderM;



                        //Agregar datos a la tabla
                        dataGridView1.Rows.Add(row);
                    }
                }
                catch (Exception es)
                {
                    MessageBox.Show("Error en conexion o creacion de tabla (Consultar con Sistemas)" + es.Message);
                }
            }

        }

        // Funcion de boton (Facturas, Entrada merc, pedidos)
        private void button3_Click(object sender, EventArgs e)
        {
            factura = textBox1.Text;
            if (factura != "")
            {
                print.Visible = true;
                dataGridView1.Visible = true;

                ((DataGridViewTextBoxColumn)dataGridView1.Columns[2]).MaxInputLength = 37;
                ((DataGridViewTextBoxColumn)dataGridView1.Columns[3]).MaxInputLength = 31;

                dataGridView1.Rows.Clear();

                try
                {
                    String str = "server=SERVER; database=SBO_TRANSMISIONES;UID=consultor;password=Consult0r*";
                    String qry = "SELECT [SBO_TRANSMISIONES].[dbo].PCH1.Quantity, " +
                        "[SBO_TRANSMISIONES].[dbo].PCH1.ItemCode," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.ItemName," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.FrgnName," +
                        "[SBO_TRANSMISIONES].[dbo].OITB.ItmsGrpNam," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.OrdrMulti," +
                        "[SBO_TRANSMISIONES].[dbo].OPCH.NumAtCard," +
                        "[SBO_TRANSMISIONES].[dbo].OITM.SuppCatNum " +
                        "FROM PCH1 INNER JOIN OPCH ON PCH1.DocEntry = OPCH.DocEntry INNER JOIN OITM ON PCH1.ItemCode = OITM.ItemCode INNER JOIN OITB ON OITM.ItmsGrpCod= OITB.ItmsGrpCod WHERE OPCH.NumAtCard ='" + factura + "'";
                    SqlConnection con = new SqlConnection(str);
                    SqlCommand cmd = new SqlCommand(qry, con);
                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();


                    while (reader.Read())
                    {
                        canti = Convert.ToInt32(reader.GetDecimal(0));
                        Codex = reader.GetString(1);
                        d1 = reader.GetString(2);
                        d2 = reader.GetString(3);
                        lane = reader.GetString(4);
                        OrderM = Convert.ToInt32(reader.GetDecimal(5));
                        if (OrderM == 0)
                            OrderM = 1;
                        factura = reader.GetString(6);

                        

                        //OPBTENER PIN SI SUPPCATNUM IS NULL
                        if (reader.IsDBNull(7))
                        {
                            Pin = Codex;

                        }
                        else
                        {
                            Pin = reader.GetString(7);
                        }

                        //Origen
                        if (Codex.IndexOf("-") != -1)
                        {
                            if (Codex.IndexOf("@") != -1)
                            {
                                origen = Codex.Substring(Codex.IndexOf("@"));
                                Codex = Codex.Substring(0, Codex.IndexOf("@"));
                            }
                            else
                            {
                                origen = Codex.Substring(Codex.IndexOf("-") + 1);
                                Codex = Codex.Substring(0, Codex.IndexOf("-"));
                            }
                        }
                        else
                        {
                            origen = "0";
                        }
                        //LIMITAR DESCRIPCION POR DEFAULT
                        if (d1.Length > 38)
                        {
                            d1 = d1.Substring(0, 38);

                        }

                        if (d2.Length > 31)
                        {
                            d2 = d2.Substring(0, 31);

                        }


                        DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                        row.Cells[0].Value = canti;
                        row.Cells[1].Value = Codex;
                        row.Cells[2].Value = d1;
                        row.Cells[3].Value = d2;
                        row.Cells[4].Value = lane;
                        row.Cells[5].Value = origen;
                        row.Cells[7].Value = Pin;
                        row.Cells[8].Value = factura;
                        row.Cells[9].Value = OrderM;



                        //Agregar datos a la tabla
                        dataGridView1.Rows.Add(row);
                    }
                }
                catch (Exception es)
                {
                    MessageBox.Show("Error en conexion o creacion de tabla (Consultar con Sistemas)" + es.Message);
                }
            }
        }
        // Funcion de boton que manda a imprimir
        private void print_Click(object sender, EventArgs e)
        {
            //Ciclo de repeticion para recorrer tabla de aplicacion 
            foreach (DataGridViewRow rowF in dataGridView1.Rows)
            {
                // rows += 1;
                //Compara si el campo de imprimir esta palomeado (true)
                if (Convert.ToBoolean(rowF.Cells["imp"].Value) == true)
                {

                    ca = Convert.ToString(rowF.Cells[0].Value);
                    Codigo = Convert.ToString(rowF.Cells[1].Value);
                    descr1 = Convert.ToString(rowF.Cells[2].Value);
                    Descr2 = Convert.ToString(rowF.Cells[3].Value);
                    lna = Convert.ToString(rowF.Cells[4].Value);
                    or = Convert.ToString(rowF.Cells[5].Value);
                    p = Convert.ToString(rowF.Cells[7].Value);
                    Ord = Convert.ToString(rowF.Cells[9].Value);
                    f = Convert.ToString(rowF.Cells[8].Value);


                    zpl = new string[9];
                    zpl[0] = ca;
                    zpl[1] = Codigo;
                    zpl[2] = descr1;
                    zpl[3] = Descr2;
                    zpl[4] = lna;
                    zpl[5] = or;
                    zpl[6] = p;
                    zpl[8] = Ord;
                    zpl[7] = f;

                    toSend = generarZPL(zpl);


                     cEtiquetas = Convert.ToInt32(ca) / Convert.ToInt32(Ord);
                    
                    

                   for (int i = 1; i <= cEtiquetas; i++)
                   {
                        sendToPrint(toSend);

                    }

                }



            }
            MessageBox.Show("Se imprimieron correctamente las etiquetas");
        }


       

        
    }
    
}
