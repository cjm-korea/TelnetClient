using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TelnetClient
{
    public partial class Form1 : Form
    {
        Excel.Application excel = new Excel.Application();
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            // Connect to the telnet server
            TcpClient client = new TcpClient("172.30.1.99", 1234);
            NetworkStream stream = client.GetStream();

            // Send the command
            byte[] command = Encoding.ASCII.GetBytes("command\r\n");
            stream.Write(command, 0, command.Length);

            // Read the response
            byte[] response = new byte[1024];
            int bytesRead = stream.Read(response, 0, response.Length);
            string responseData = Encoding.ASCII.GetString(response, 0, bytesRead);

            rtbResponse.AppendText(responseData);
            // Display the response
            MessageBox.Show(responseData);

            // Close the connection
            stream.Close();
            client.Close();
        }

        private void xlsxBtn_Click(object sender, EventArgs e)
        {
            workbook = excel.Workbooks.Add();
            worksheet = workbook.ActiveSheet;

            // 엑셀 파일에 데이터 쓰기
            worksheet.Cells[1, 1] = rtbResponse.Text;

            // 엑셀 파일 저장
            workbook.SaveAs("C:\\Temp\\test.xlsx");
            workbook.Close();
            excel.Quit();

            MessageBox.Show("Data saved to Excel file.");
        }
    }
}
