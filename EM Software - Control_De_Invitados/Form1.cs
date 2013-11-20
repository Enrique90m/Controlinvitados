using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.SqlServerCe;
using System.Windows.Forms.DataVisualization.Charting;
using System.Net.Mail;
using System.Net.Mime;

namespace EM_Software___Control_De_Invitados
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //CUANDO ARRANCA EL SISTEMA, CONECTA CON LA BD           
           // this.conectaBD();                           
        }

        public DataTable dt;
        public SqlCeDataAdapter da;
        public SqlCeConnection PathBD;
        public DataRow row;
        string valor;
        int cambios = 0;
        bool enviarcorreo = true;

        #region Metodos
        public void conectaBD()
        {         
           //busca la path de la aplicacion y le agrega la base de datos en string            
            string path = "C:\\EM Software - Control_De_Invitados\\EM Software - Control_De_Invitados\\BD\\Control_de_invitados_BD.sdf;Persist Security Info=False;";
            //             C:\EM Software - Control_De_Invitados\EM Software - Control_De_Invitados\BD
            SqlCeConnection PathBD = new SqlCeConnection("Data source="+path);
            //abre la conexion
            try
            {
                string SeleccionaTodosLosDatos = "SELECT * FROM Invitados ORDER BY Num_invitado";
                da = new SqlCeDataAdapter(SeleccionaTodosLosDatos, PathBD);

                // Crear los coma;ndos de insertar, actualizar y eliminar
                SqlCeCommandBuilder cb = new SqlCeCommandBuilder(da);
                // Asignar los comandos al DataAdapter
                // (se supone que lo hace automáticamente, pero...)
                da.UpdateCommand = cb.GetUpdateCommand();
                da.InsertCommand = cb.GetInsertCommand();
                da.DeleteCommand = cb.GetDeleteCommand();
                dt = new DataTable();
                // Llenar la tabla con los datos indicados
                da.Fill(dt);

                PathBD.Open();
            }
            catch (Exception w)
            {
                MessageBox.Show(w.ToString());                           
            }
         
        }
        private void BuscarAlgoenTabla(string comando)
        {

            string path = "C:\\EM Software - Control_De_Invitados\\EM Software - Control_De_Invitados\\BD\\Control_de_invitados_BD.sdf;Persist Security Info=False;";
            SqlCeConnection PathBD = new SqlCeConnection("Data source=" + path);
            //abre la conexion
            try
            {
                da = new SqlCeDataAdapter(comando, PathBD);

                // Crear los comandos de insertar, actualizar y eliminar
                SqlCeCommandBuilder cb = new SqlCeCommandBuilder(da);                
                dt = new DataTable();
                // Llenar la tabla con los datos indicados
                da.Fill(dt);

                PathBD.Open();
            }
            catch (Exception w)
            {
                MessageBox.Show(w.ToString());
                return;
            }
        }
        private bool PoneDatosEnRenglon(bool error)
        {

            foreach (Control c in this.DatosInvitado_groupBox1.Controls)
                if (c is TextBox || c is MaskedTextBox)
                    if (c.Name != "textBox1" && c.Name != "textBox2")
                        if (c.Text == "" || c.Text == null || string.IsNullOrWhiteSpace(c.Text))
                        {   //FALTO CAPTURAR ALGO. LE AVISA AL USUARIO Y SE SALE
                            MessageBox.Show("Falto capturar datos, porfavor verifique!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            error = true;
                            return (error);
                        }           


            //VERIFICA QUE HAYA INGRESADO CORREO E IMAGEN
            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrWhiteSpace(textBox1.Text))
            {
                DialogResult respuesta = MessageBox.Show("No capturo correo electronico, desea continuar? ", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (respuesta.Equals(DialogResult.No))
                {                   
                    return (error=true);
                }
                else
                    enviarcorreo = false; 
            }

            if (string.IsNullOrEmpty(textBox2.Text) || string.IsNullOrWhiteSpace(textBox2.Text))
            {
                DialogResult respuesta = MessageBox.Show("No capturo imagen, desea continuar? ", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (respuesta.Equals(DialogResult.No))
                    return (error=true);
            }
           
            //crea nuevo renglon y lo llena con los datos
            row = dt.NewRow();
            DateTime x = DateTime.Now;            
            row["Num_invitado"] = dt.Rows.Count+1;
            row["Nombre"] = Nombre_datosinvidatoTextBox.Text;
            row["ApellidoP"] = ApellidoP_datosinvidatoTextBox.Text;
            row["ApellidoM"] = ApellidoM_datosinvidatoTextBox.Text;
            row["Invitados"] = Invitados_datosinvitadoMasked.Text;
            row["Asistio"] = 0;
            row["NumMesa"] = NumMesa_datosinvitadoMasked.Text;
            row["FechaRegistro"] = x.ToShortTimeString();
            row["correo"] = textBox1.Text;        
            return (error);
        }
        private void AgregaABD()
        {
            //agrega renglon a la tabla virtual
            dt.Rows.Add(row);
            //la agrega a la base de datos fisica
            try
            {
                da.Update(dt);
                dt.AcceptChanges();
            }
            catch (DBConcurrencyException ex)
            {
                MessageBox.Show("Error de concurrencia:\n" + ex.Message);
                return;
            }
            MessageBox.Show("Invitado agregado correctamente!");
        }
        private void BuscarUnInvitado(string celda, string Dato)//celda es la celda en la que buscara , solo puede ser : Nombre o ApellidoP
        {
            bool EncontroAlgo = false;
            int NumCelda = 0;
            //Si busca por nombre pone numcelda=1, si no la pone a 2
            if (celda == "Nombre")
                NumCelda = 1;
            else
                NumCelda = 2;
            //busca en todos los rows del datagrid 
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                if (dataGridView1.Rows[i].Cells[NumCelda].Value.ToString().Contains(Dato))//si la celda contiene a dato , osea lo encontro
                {
                    DataGridViewRow row = dataGridView1.Rows[i];//guarda el renglon
                    dataGridView1.DataSource = row;
                    EncontroAlgo = true;
                }

            //si encontro algo = false no encontro nada encontes:
            if (EncontroAlgo == false)
                MessageBox.Show("No se encontro ningun registro parecido al dato ingresado, porfavor verifique!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

            return;
        }       
        private void saca_estadisticas()
        {
            bool error = false;
            //SACO TOTAL DE INVITADOS Y LO PONGO EN TEXTBOX           
            this.BuscarAlgoenTabla("SELECT sum(Invitados) as suma FROM Invitados");
            Estadisticas_TotInvitados_TB.Text = dt.Rows[0]["suma"].ToString();

            //SUMO LOS INVITADOS QUE ASISTIERON 
            this.BuscarAlgoenTabla("SELECT sum(InvitadosQAsistieron) as suma FROM Invitados WHERE Asistio = 1");
            string algo = dt.Rows[0]["suma"].ToString();
            if (dt.Rows[0]["suma"].ToString() == "0" || dt.Rows[0]["suma"].ToString() == "")
            {
                error = true;
                Estadisticas_InitQAsist.Text = "0";
            }
            else
            {
                Estadisticas_InitQAsist.Text = algo.ToString();//dt.Rows[0]["suma"].ToString();
            }

            //SACA TOTAL DE LOS QUE NO ASISIERON
            this.BuscarAlgoenTabla("SELECT sum(Invitados) as suma FROM Invitados WHERE Asistio = 0");
            if (dt.Rows[0]["suma"].ToString() == "0" || dt.Rows[0]["suma"].ToString() == "")
            {
                //long invitadosQnoAsist = long.Parse(dt.Rows[0]["suma"].ToString());
                this.BuscarAlgoenTabla("SELECT  * FROM Invitados WHERE Asistio = 1");
                //invitadosQnoAsist = invitadosQnoAsist + long.Parse(dt.Rows[0]["RESULTADO"].ToString());
                long total = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["Invitados"].ToString() == "")
                        dt.Rows[i]["Invitados"] = 0;

                    long invitados = long.Parse(dt.Rows[i]["Invitados"].ToString());

                    if (dt.Rows[i]["InvitadosQAsistieron"].ToString() == "")
                        dt.Rows[i]["InvitadosQAsistieron"] = "0";

                    long invitadosQAsistieron = long.Parse(dt.Rows[i]["InvitadosQAsistieron"].ToString());
                   
                    total = total + (invitados - invitadosQAsistieron);
                }
                Estadisticas_InvQNoAsist.Text = total.ToString();
            }
            else
            {
                long invitadosQnoAsist = long.Parse(dt.Rows[0]["suma"].ToString());
                this.BuscarAlgoenTabla("SELECT  * FROM Invitados WHERE Asistio = 1");
                if (dt.Rows.Count == 0)
                    Estadisticas_InvQNoAsist.Text = invitadosQnoAsist.ToString();
                else
                {
                    long total = 0;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["Invitados"].ToString() == "")
                            dt.Rows[i]["Invitados"] = 0;

                        long invitados = long.Parse(dt.Rows[i]["Invitados"].ToString());

                        if (dt.Rows[i]["InvitadosQAsistieron"].ToString() == "")
                            dt.Rows[i]["InvitadosQAsistieron"] = "0";

                        long invitadosQAsistieron = long.Parse(dt.Rows[i]["InvitadosQAsistieron"].ToString());

                        total = total + (invitados - invitadosQAsistieron);
                    }
                    total = total + invitadosQnoAsist;
                    Estadisticas_InvQNoAsist.Text = total.ToString();
                }
            }


            //NO HAY NADIE CON ASISTENCIA, PONE CERO Y SE SALE
            if (error == true)
            {
                Estadisticas_Promedio.Text = "0%";
            }
            else
            {
                //PROMEDIO DE ASISTENCIA
                double promedio = double.Parse(Estadisticas_InitQAsist.Text.ToString()) / double.Parse(Estadisticas_TotInvitados_TB.Text.ToString());
                promedio = promedio * 100;
                promedio = System.Math.Round(promedio, 5);
                Estadisticas_Promedio.Text = promedio.ToString() + "%";
            }

            //chart1.

            int invitadosQasistieron = int.Parse(Estadisticas_InitQAsist.Text);
            int invitadosQNoAsist = int.Parse(Estadisticas_InvQNoAsist.Text);

            chart1.Series["Invitados Que Asistieron"].Points.Clear();
            chart1.Series["Invitados Que No Asistieron"].Points.Clear();

            //chart1.Series["Invitados Que Asistieron"].Points[0].BorderW
            chart1.Series["Invitados Que Asistieron"].Points.AddXY(1, invitadosQasistieron);
            chart1.Series["Invitados Que No Asistieron"].Points.AddXY(2, invitadosQNoAsist);

            //  chart1.Series["Series1"].Points = 3
            // chart1.Series["Series2"].YValueMembers = "Y";               


            //REGRESA LOS DATOS A SU ORIGEN
            this.BuscarAlgoenTabla("SELECT * FROM Invitados ORDER BY Num_Invitado");
        }        
        #endregion

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            //pone la pagina correspondiente y desabilita los controles exepto los labels
            this.tabControl1.SelectTab(EditarInvitado);
            foreach (Control control in this.Resultados_editarGB.Controls)
                if (control is TextBox || control is MaskedTextBox || control is Button)
                    control.Enabled = false;
        }

        private void BuscarInvitado_tabcontrolboton_Click(object sender, EventArgs e)
        {
            // si ya esta seleccionada no hace nada
            if (this.tabControl1.SelectedTab == Estadisticas_TP)
                return;
           
            //primero verifica que haya registros , si NO hay, NO abre la ventana
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("No hay registros en la tabla invitados!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
                       
            this.tabControl1.SelectTab(Estadisticas_TP);

            //CAMBIOS=0 - PRIMERA VEZ QUE EJECUTA EL CODIGO DE ESTE METODO
            //CAMBIOS=1 - NO HUBO CAMBIOS, NO EJECUTA EL METODO
            //CAMBIOS=2 - HUBO CAMBIO, EJECUTA EL CODIGO Y LO REGRESA A 1
            if (cambios == 0)
            {
                this.saca_estadisticas();
                cambios = 1;
                this.BuscarAlgoenTabla("SELECT * FROM Invitados");
            }
            else
                if (cambios == 1)
                    return;
                else
                    if (cambios == 2)
                    {
                        this.saca_estadisticas();
                        cambios = 1;
                        this.BuscarAlgoenTabla("SELECT * FROM Invitados");
                    }
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            this.tabControl1.SelectTab(AgregarInvitado);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //limpia los datos si el control es un textbox
            foreach (Control control in this.DatosInvitado_groupBox1.Controls)
                if (control is TextBox || control is MaskedTextBox)                   
                    control.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool error = false;          
            error = this.PoneDatosEnRenglon(error);          
            //SI AL VALIDAR HUBO ERROR, NO LO AGREGA
           if (error == true)
               return;
           
               this.AgregaABD();
               cambios = 2;

               if (enviarcorreo == false)
                   return;

               
                   DialogResult respuesta = MessageBox.Show("Desea enviar correo electronico a su invitado? ", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                   if (respuesta.Equals(DialogResult.No))                                     
                       return;
                   


           MailMessage mail = new MailMessage();
           SmtpClient SmtpServer = new SmtpClient();
           SmtpServer.Credentials = new System.Net.NetworkCredential("enviainvitacion9@gmail.com", "enviainvitacion");
           SmtpServer.Port = 587;
           SmtpServer.Host = "smtp.gmail.com";
           SmtpServer.EnableSsl = true;           
  
           try
           {
               mail.From = new MailAddress(textBox1.Text, "Invitacion", System.Text.Encoding.UTF8);
               mail.To.Add(textBox1.Text);
               mail.Subject = "";
               mail.Body = "Buen dia Sr/Sra. " + Nombre_datosinvidatoTextBox.Text + "" + ApellidoP_datosinvidatoTextBox + "\n\n" +
                           "Usted esta cordialmente invitado a nuestra fiesta";

               if (textBox2.Text != "")
               {
                   LinkedResource logo = new LinkedResource(textBox2.Text);
                   logo.ContentId = "Logo";
                   string htmlview;
                   htmlview = "<html><body><table border=2><tr width=100%><td><img src=cid:Logo alt=companyname /></td><td></td></tr></table><hr/></body></html>";
                   AlternateView alternateView1 = AlternateView.CreateAlternateViewFromString(htmlview + "", null, MediaTypeNames.Text.Html);
                   alternateView1.LinkedResources.Add(logo);
                   mail.AlternateViews.Add(alternateView1);
               }
               mail.IsBodyHtml = true;
               mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
               //mail.ReplyTo = new MailAddress(TextBox1.Text);
              // SmtpServer.Send(mail);
               //mail.ReplyTo = new MailAddress(TextBox1.Text);
               SmtpServer.Send(mail);
           }
           catch (Exception ex)
           {
               MessageBox.Show("No se pudo enviar correo electronico al invitado, le presentamos algunas sugerencias : \n \n-Verifique que tenga conexion a internet \n-Verifique que el correo electronico ingresado sea el correcto \n-Actualize el certificado de la pagina ","",MessageBoxButtons.OK,MessageBoxIcon.Warning);
               MessageBox.Show("No se envio correo electronico!","",MessageBoxButtons.OK,MessageBoxIcon.Warning);
               return;
           }

           MessageBox.Show("Se agrego y se envio el correo al invitado correctamente");
        }       

        private void Nombre_datosinvidatoTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar) && e.KeyChar != '\b')
                e.Handled = true;
        }

        private void ApellidoP_datosinvidatoTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar) && e.KeyChar != '\b')
                e.Handled = true;
        }

        private void ApellidoM_datosinvidatoTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar) && e.KeyChar != '\b')
                e.Handled = true;
        }               
          

        private void Editar_Click(object sender, EventArgs e)
        {           
            try
            {
                da.Update(dt);
                dt.AcceptChanges();
            }
            catch (DBConcurrencyException ex)
            {
                MessageBox.Show("Error de concurrencia:\n" + ex.Message);
                return;
            }
            MessageBox.Show("Cambios guardados correctamente");
            cambios = 2;
            return;
             
        }
        
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            //guardo el dato antes de que lo modifique
            valor = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

        }

        private void DatoABuscar_KeyUp(object sender, KeyEventArgs e)
        {
            if (Nombre_RadioBttn.Checked)
            {
                DataView dataView = dt.DefaultView;              
                dataView.RowFilter = string.Format("Nombre LIKE '{0}%'", DatoABuscar.Text);
                dataGridView1.DataSource = dataView;
            }
            else
            if(ApellidoP_RadioBttn.Checked)
            {
                DataView dataView = dt.DefaultView;
                dataView.RowFilter = string.Format("ApellidoP LIKE '{0}%'", DatoABuscar.Text);
                dataGridView1.DataSource = dataView;
            }
            else
               if(ApellidoM.Checked)
            {
                DataView dataView = dt.DefaultView;               
                dataView.RowFilter = string.Format("ApellidoM LIKE '{0}%'", DatoABuscar.Text);
                dataGridView1.DataSource = dataView;
            }
               else
               {
                   //si esta vacio datoabuscar, rellena el datagridview
                   if (DatoABuscar.Text == "" || DatoABuscar == null)
                   {
                       DataView dataView1 = dt.DefaultView;
                       dataView1.RowFilter = string.Format("ApellidoM LIKE '{0}%'", DatoABuscar.Text);
                       dataGridView1.DataSource = dataView1;
                       return;
                   }                                    
               }
        }      

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {           
                e.Control.KeyPress -= new KeyPressEventHandler(Valid);
                e.Control.KeyPress += new KeyPressEventHandler(Valid);
            
        }
     
        private void Valid(object sender, KeyPressEventArgs e)
        {
            //si esta vacio se regresa
            if (e.KeyChar==8 && dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString() == "")
            {
                e.Handled = true;
                return;
            }
             //si la celda donde oprimio el boton es de tipo entero, verifica que no haya ingresado letras o caracteres especiales
             string headerText = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name;
             if (headerText.Equals("Num_invitado") || headerText.Equals("Invitados") || headerText.Equals("InvitadosQAsistieron") || headerText.Equals("NumMesa"))
             {
                if (e.KeyChar >= 48 && e.KeyChar <= 57/*Admite los numeros del 0 al 9*/|| e.KeyChar == 8/* codigo ascii del backspace*/)
                      e.Handled = false;
                     else 
                    e.Handled = true;
                 return;   
             }

             //si la celda donde oprimio el boton es de tipo string, verifica que no haya ingresado numeros
             if (headerText.Equals("Nombre") || headerText.Equals("ApellidoP") || headerText.Equals("ApellidoM"))
             {
                 if ((e.KeyChar >= 65 && e.KeyChar <= 90) || (e.KeyChar >= 97 && e.KeyChar <= 122) || e.KeyChar == 8 || e.KeyChar == 'ñ' || e.KeyChar == 'Ñ')
                     e.Handled = false;
                 else
                     e.Handled = true;
             }
             
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //si esta vacio se regresa
            if (dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString() == "" || dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString() == null)
            {
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[dataGridView1.CurrentCell.ColumnIndex].Value = valor;              
                return;
            }

        }       

        private void toolStripButton2_Click(object sender, EventArgs e)
        {           
            this.tabControl1.SelectTab(EditarInvitado);
            this.dataGridView1.DataSource = dt;
        }

        private void Invitados_datosinvitadoMasked_Click(object sender, EventArgs e)
        {
            Invitados_datosinvitadoMasked.Select(0, 0);
        }

        private void SalirSistema_Click(object sender, EventArgs e)
        {
            DialogResult respuesta = MessageBox.Show("Esta seguro que desea salir? ", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (respuesta.Equals(DialogResult.Yes))
                Application.Exit();            
        }

        private void BorraDatosBuscar_Click(object sender, EventArgs e)
        {
            DatoABuscar.Text = "";
        }        

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text.EndsWith("@hotmail.com") || textBox1.Text.EndsWith("@gmail.com") || textBox1.Text.EndsWith("@yahoo.com"))
            {                     
                return;
            }
            else
            {
                MessageBox.Show("El correo ingresado no es valido, debe terminar en : @hotmail.com o @gmail.com o @yahoo.com, Porfavor verifique!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox1.Text = "";
                return;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ruta = new OpenFileDialog();
            if (ruta.ShowDialog() == DialogResult.OK)
                if (ruta.FileName.Contains(".jpg") || ruta.FileName.Contains(".gif") || ruta.FileName.Contains(".png") || ruta.FileName.Contains(".JPG") || ruta.FileName.Contains(".Jpg"))
                textBox2.Text = ruta.FileName;
                else
                    MessageBox.Show("El archivo seleccionado no es una imagen", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
        }

        private void Invitados_datosinvitadoMasked_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar)  && e.KeyChar != '\b')
                e.Handled = true;
        }

        private void Invitados_datosinvitadoMasked_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
                e.Handled = true;
        }

        private void NumMesa_datosinvitadoMasked_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
                e.Handled = true;
        }

        private void Control_de_invitados_Load(object sender, EventArgs e)
        {          
            try
            {
                string path = "C:\\EM Software - Control_De_Invitados\\EM Software - Control_De_Invitados\\BD\\Control_de_invitados_BD.sdf;Persist Security Info=False;";
                SqlCeConnection PathBD = new SqlCeConnection("Data source=" + path);
                PathBD.Open();
            }
            catch (Exception c)
            {
              
                MessageBox.Show(c.ToString());
                Application.Exit();
                return;
            }
            this.conectaBD();        
        }
        
                            
     }
}