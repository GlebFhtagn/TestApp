﻿using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Linq;

namespace TestApp
{
    public partial class ComponentForm : Form
    {
        Word._Application application;
        Word._Document document;

        Object templatePathObj = Application.StartupPath + "\\DataBase_Report.docx";
        
        private void fillComponentTree()
        {
            ComponentsTree.Nodes.Clear();
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
            {
                connection.Open();

                String command = "SELECT ID, NAME " +
                                    "FROM TOPCOMP ";

                SqlCommand sc = new SqlCommand(command, connection);
                SqlDataReader reader = sc.ExecuteReader();

                while (reader.Read())
                {

                    int topComponentId = (int)reader.GetValue(0);
                    String topComponentName = reader.GetValue(1).ToString();

                    TreeNode topNode = new TreeNode(topComponentName);
                    ComponentsTree.Nodes.Add(topNode);

                    PopulateTopNode(topComponentId, topNode);
                }

                connection.Close();
            }
        }

        public ComponentForm()
        {
            InitializeComponent();
            fillComponentTree();
        }

        private void PopulateTopNode(int TID, TreeNode topNode)
        {
            String command =    "SELECT INCOMP.ID, NAME, QUANTITY " +
                                "FROM INCOMP INNER JOIN TOP_IN_COMP " +
                                "ON INCOMP.ID = TOP_IN_COMP.ID " +
                                "AND TOP_IN_COMP.TID = " + TID;

            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["PopulateConnection"].ConnectionString))
            {
                connection.Open();

                SqlCommand sc = new SqlCommand(command, connection);
                SqlDataReader currReader = sc.ExecuteReader();

                while (currReader.Read())
                {
                    int componentId = (int) currReader.GetValue(0);
                    String componentName = currReader.GetValue(1).ToString();
                    String componentQuantity = currReader.GetValue(2).ToString();

                    TreeNode node = new TreeNode(componentName + " (" + componentQuantity + ")");
                    topNode.Nodes.Add(node);

                    PopulateNode(componentId, node);
                }
                connection.Close();
            }
        }

        private void PopulateNode(int ID, TreeNode parentNode)
        {
            String command =    "SELECT INCOMP.ID, NAME, QUANTITY " +
                                "FROM INCOMP INNER JOIN IN_IN_COMP " +
                                "ON INCOMP.ID = IN_IN_COMP.ID " +
                                "AND IN_IN_COMP.PID = " + ID;

            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["PopulateConnection"].ConnectionString))
            {
                connection.Open();

                SqlCommand sc = new SqlCommand(command, connection);
                SqlDataReader currReader = sc.ExecuteReader();

                while (currReader.Read())
                {
                    int componentId = (int) currReader.GetValue(0);
                    String componentName = currReader.GetValue(1).ToString();
                    int componentQuantity = (int) currReader.GetValue(2);

                    TreeNode node = new TreeNode(componentName + " (" + componentQuantity+ ")");
                    parentNode.Nodes.Add(node);

                    PopulateNode(componentId, node);
                }
                connection.Close();
            }
        }

        private void КомпонентToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddTopForm addTopForm = new AddTopForm("");

            WrongName:
            try
            {
                if (addTopForm.ShowDialog(this) == DialogResult.OK)
                {
                    String componentName = addTopForm.GetComponent();

                    TreeNodeCollection rootNodes = ComponentsTree.Nodes;

                    bool containsFlag = false;
                    containsFlag = isContainsInDB(componentName);

                    if (!containsFlag)
                    {
                        TreeNode topNode = new TreeNode(componentName);
                        ComponentsTree.Nodes.Add(topNode);
                        using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                        {
                            connection.Open();
                            new SqlCommand("INSERT INTO TOPCOMP (NAME) VALUES ('" + addTopForm.GetComponent() + "')", connection).ExecuteNonQuery();
                            connection.Close();
                        }

                        addTopForm.Close();
                    }
                    else
                    {
                        addTopForm.containAllert(componentName);
                        addTopForm.Hide();
                        goto WrongName;
                    }
                }
            }
            catch(Exception error)
            {
                addTopForm.changeStatus(error.Message);
                addTopForm.Hide();
                goto WrongName;
            }

        }

        private void вложенныйКомпонентToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddForm addForm = new AddForm();
            WrongName:
            if (addForm.ShowDialog(this) == DialogResult.OK)
            {
                try
                {
                    TreeNode parentNode = ComponentsTree.SelectedNode;
                    String parentName = returnNameOfTheNode(parentNode);
                    int parentQuantity = returnQuantityOfTheNode(parentNode);
                    String componentName = addForm.GetComponent();
                    int componentQuantity = addForm.GetQuantity();

                    if (checkForRecursion(parentNode, componentName))
                    {
                        using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                        {
                            String command = "";
                            SqlCommand sqlCommand = null;
                            if (!isContainsInDB(componentName))
                            {
                                connection.Open();
                                command = "INSERT INTO INCOMP (NAME) VALUES ('" + componentName + "')";
                                sqlCommand = new SqlCommand(command, connection);
                                sqlCommand.ExecuteNonQuery();
                                connection.Close();
                            }

                            if (parentNode.Parent == null)
                            {
                                int parentId = getNodeIdByName(parentName);
                                int nodeId = getNodeIdByName(componentName);

                                command = "SELECT * FROM TOP_IN_COMP WHERE TID = " + parentId + " AND ID = " + nodeId;
                                sqlCommand = new SqlCommand(command, connection);
                                connection.Open();
                                SqlDataReader reader = sqlCommand.ExecuteReader();
                                if (reader.HasRows)
                                {
                                    reader.Close();
                                    addForm.containAllert(componentName, componentQuantity);
                                    addForm.Hide();
                                    goto WrongName;
                                }
                                else
                                {
                                    reader.Close();
                                    command = "INSERT INTO TOP_IN_COMP (TID, ID, QUANTITY) VALUES (" + parentId + ", " + nodeId + ", " + componentQuantity + ")";
                                    sqlCommand = new SqlCommand(command, connection);
                                    sqlCommand.ExecuteNonQuery();
                                }
                                connection.Close();
                            }
                            else
                            {
                                int parentId = getNodeIdByName(parentName);
                                int nodeId = getNodeIdByName(componentName);

                                command = "SELECT * FROM IN_IN_COMP WHERE PID = " + parentId + " AND ID = " + nodeId;
                                sqlCommand = new SqlCommand(command, connection);
                                connection.Open();
                                SqlDataReader reader = sqlCommand.ExecuteReader();
                                bool relationExists = reader.HasRows;
                                reader.Close();
                                if (relationExists)
                                {
                                    addForm.containAllert(componentName, componentQuantity);
                                    addForm.Hide();
                                    goto WrongName;
                                }
                                else
                                {
                                    command = "INSERT INTO IN_IN_COMP (PID, ID, QUANTITY) VALUES (" + parentId + ", " + nodeId + ", " + componentQuantity + ")";
                                    sqlCommand = new SqlCommand(command, connection);
                                    sqlCommand.ExecuteNonQuery();
                                }
                            }
                            connection.Close();
                        }
                    }
                    else
                    {
                        addForm.recursionFound(componentName, componentQuantity);
                        addForm.Hide();
                        goto WrongName;
                    }
                    addForm.Close();
                    fillComponentTree();
                }
                catch(Exception error)
                {
                    addForm.changeStatus(error.Message);
                    addForm.Hide();
                    goto WrongName;
                }
            }
        }
        
        private void переименоватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = ComponentsTree.SelectedNode;
            TreeNode parentNode = node.Parent;
            int nodeId, nodeQuantity, parentId, parentQuantity;
            String nodeName, parentName;

            if (parentNode == null)
            {
                nodeId = getNodeId(node);
                AddTopForm addTopForm = new AddTopForm(node.Text);
                WrongName1:
                try
                {
                    if (addTopForm.ShowDialog(this) == DialogResult.OK)
                    {
                        String componentName = addTopForm.GetComponent();
                        if (checkForRecursion(parentNode, componentName) && findNodeByName(node.Nodes, componentName).Length == 0)
                            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                            {
                                connection.Open();
                                String command = "UPDATE TOPCOMP SET NAME = '" + componentName + "' WHERE NAME = '" + node.Text + "'";
                                SqlCommand sc = new SqlCommand(command, connection);
                                sc.ExecuteNonQuery();
                                node.Text = componentName;
                            }
                        else
                        {
                            addTopForm.recursionFound(componentName);
                            addTopForm.Hide();
                            goto WrongName1;
                        }
                    }
                }
                catch (Exception error)
                {
                    addTopForm.changeStatus(error.Message);
                    addTopForm.Hide();
                    goto WrongName1;
                }
            }
            else
            {
                nodeId = getNodeId(node);
                nodeName = returnNameOfTheNode(node);
                nodeQuantity = returnQuantityOfTheNode(node);
                parentId = getParentId(node);
                parentName = returnNameOfTheNode(parentNode);
                parentQuantity = returnQuantityOfTheNode(parentNode);

                AddForm addForm = new AddForm(nodeName, nodeQuantity);
                WrongName2:
                if (addForm.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        String componentName = addForm.GetComponent();
                        int componentQuantity = addForm.GetQuantity();
                        if (checkForRecursion(node.Parent, componentName) && findNodeByName(node.Nodes, componentName).Length == 0)
                        {
                            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                            {
                                connection.Open();
                                int id = getNodeId(node);

                                if (!isContainsInDB(componentName))
                                {
                                    String command = "UPDATE INCOMP SET NAME = '" + componentName + "' WHERE ID = " + id;
                                    SqlCommand sc = new SqlCommand(command, connection);
                                    sc.ExecuteNonQuery();

                                    command = "UPDATE IN_IN_COMP SET QUANTITY = " + componentQuantity + " WHERE ID = " + id + " AND PID = " + parentId;
                                    sc = new SqlCommand(command, connection);
                                    sc.ExecuteNonQuery();
                                }
                                else
                                {
                                    if (parentNode.Parent == null)
                                    {
                                        int newId = getNodeIdByName(componentName);
                                        String command = "DELETE TOP_IN_COMP WHERE ID = " + id + " AND TID = " + parentId;
                                        SqlCommand sc = new SqlCommand(command, connection);
                                        sc.ExecuteNonQuery();

                                        command = "INSERT INTO TOP_IN_COMP (TID, ID, QUANTITY) VALUES (" + parentId + ", " + newId + ", " + componentQuantity + ")";
                                        sc = new SqlCommand(command, connection);
                                        sc.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        int newId = getNodeIdByName(componentName);
                                        String command = "DELETE IN_IN_COMP WHERE ID = " + id + " AND PID = " + parentId;
                                        SqlCommand sc = new SqlCommand(command, connection);
                                        sc.ExecuteNonQuery();

                                        command = "INSERT INTO IN_IN_COMP (PID, ID, QUANTITY) VALUES (" + parentId + ", " + newId + ", " + componentQuantity + ")";
                                        sc = new SqlCommand(command, connection);
                                        sc.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                        else
                        {
                            addForm.recursionFound(componentName, componentQuantity);
                            addForm.Hide();
                            goto WrongName2;
                        }
                    }
                    catch(Exception error)
                    {
                        addForm.changeStatus(error.Message);
                        addForm.Hide();
                        goto WrongName2;
                    }
                }
            }
            fillComponentTree();
        }


        private void deleteAllRelations(TreeNode node)
        {
            String componentName = returnNameOfTheNode(node);
            int componentId = getNodeId(node);

            String command = (node.Parent != null ? "DELETE FROM IN_IN_COMP WHERE PID = " + componentId : "DELETE FROM TOP_IN_COMP WHERE PID = " + componentId);
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
            {
                SqlCommand sqlCommand = new SqlCommand(command, connection);
                sqlCommand.ExecuteNonQuery();
            }
            foreach (TreeNode childNode in node.Nodes)
            {
                deleteAllRelations(childNode);
            }
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode node = ComponentsTree.SelectedNode;
            TreeNode parentNode = node.Parent;
            if (node == null)
                return;
            if (parentNode == null)
            {
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                {
                    String command = "DELETE FROM TOPCOMP WHERE NAME = '" + node.Text + "'";
                    SqlCommand sc = new SqlCommand(command, connection);
                    connection.Open();
                    sc.ExecuteNonQuery();
                    ComponentsTree.Nodes.Remove(node);
                    connection.Close();
                }
            }
            else
            {
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                {
                    connection.Open();

                    String сomponentName = returnNameOfTheNode(node);
                    int componentId = getNodeId(node);

                    String parentName = returnNameOfTheNode(parentNode);
                    int parentId = getNodeId(parentNode);
                    if (parentNode.Parent != null)
                    {
                        String command1 = "DELETE FROM IN_IN_COMP WHERE PID = " + parentId + " AND ID = " + componentId;
                        SqlCommand sqlCommand1 = new SqlCommand(command1, connection);
                        sqlCommand1.ExecuteNonQuery();
                    }
                    else
                    {
                        String command1 = "DELETE FROM TOP_IN_COMP WHERE TID = " + parentId + " AND ID = " + componentId;
                        SqlCommand sqlCommand1 = new SqlCommand(command1, connection);
                        sqlCommand1.ExecuteNonQuery();
                    }
                    
                    String command = "SELECT ID FROM TOP_IN_COMP WHERE ID = " + componentId;
                    SqlCommand sqlCommand = new SqlCommand(command, connection);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    bool leftInTopRelations = reader.HasRows;
                    reader.Close();

                    command = "SELECT ID FROM IN_IN_COMP WHERE ID = " + componentId;
                    sqlCommand = new SqlCommand(command, connection);
                    reader = sqlCommand.ExecuteReader();
                    bool leftInElseRelations = reader.HasRows;
                    reader.Close();

                    if (!leftInTopRelations && !leftInElseRelations)
                    {
                        command = "DELETE FROM INCOMP WHERE ID =" + componentId;
                        sqlCommand = new SqlCommand(command, connection);
                        sqlCommand.ExecuteNonQuery();
                    }
                    connection.Close();
                }
                ComponentsTree.Nodes.Remove(node);
            }
        }

        private void отчетОСводномСоставеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Word Documents|*.doc; *.docx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                templatePathObj = openFileDialog1.FileName;
            }
            else return;
            Object missingObj = Missing.Value;
            Object trueObj = true;
            Object falseObj = false;            

            application = new Word.Application();


            TreeNode selectedNode = ComponentsTree.SelectedNode;
            if (selectedNode.Nodes.Count > 0)
            {
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                {
                    connection.Open();
                    String command = (selectedNode.Parent == null ?
                                    "SELECT COUNT(*) FROM TOPCOMP INNER JOIN TOP_IN_COMP on NAME = '" + returnNameOfTheNode(selectedNode) + "' and TOPCOMP.ID =  TOP_IN_COMP.TID" :
                                    "SELECT COUNT(*) FROM INCOMP INNER JOIN IN_IN_COMP ON NAME = '" +returnNameOfTheNode(selectedNode) + "' and INCOMP.ID =  IN_IN_COMP.PID");
                    SqlCommand sqlCommand = new SqlCommand(command, connection);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    reader.Read();
                    int rowsCount = (int)reader.GetValue(0);
                    reader.Close();
                    
                    try
                    {
                        application = new Word.Application();
                        document = application.Documents.Open(templatePathObj);
                    }
                    catch (Exception error)
                    {
                        document.Close(ref falseObj, ref missingObj, ref missingObj);
                        application.Quit(ref missingObj, ref missingObj, ref missingObj);
                        document = null;
                        application = null;
                        Console.WriteLine(error.Message);
                    }
                    application.Visible = true;

                    Object start = 0;
                    Object end = 0;
                    Word.Range currentRange = document.Range(ref start, ref end);
                    Word.Table table = document.Tables.Add(currentRange, rowsCount, 2, missingObj, missingObj);
                    Object style = "Table Grid";
                    Word.Styles styles = table.Range.Document.Styles;
                    table.set_Style(style);
                    Word.Range wordcellrange;
                    command = (selectedNode.Parent == null ?
                                   "SELECT INCOMP.NAME, QUANTITY FROM TOPCOMP INNER JOIN TOP_IN_COMP ON NAME = '" + returnNameOfTheNode(selectedNode) + "' and TOPCOMP.ID =  TOP_IN_COMP.TID INNER JOIN INCOMP ON INCOMP.ID = TOP_IN_COMP.ID" :
                                   "SELECT INCOMP.NAME, QUANTITY FROM INCOMP INNER JOIN IN_IN_COMP ON NAME = '" + returnNameOfTheNode(selectedNode) + "' and INCOMP.ID =  IN_IN_COMP.PID");
                    sqlCommand = new SqlCommand(command, connection);
                    reader = sqlCommand.ExecuteReader();

                    for (int m = 0; m < table.Rows.Count; m++)
                    {
                        reader.Read();
                        for (int n = 0; n < table.Columns.Count; n++)
                        {
                            wordcellrange = table.Cell(m+1, n+1).Range;
                            wordcellrange.Text = "" + reader.GetValue(n);
                        }
                    }
                        reader.Close();                            
                    
                    connection.Close();
                }

            }
        }

        private void ComponentsTree_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ComponentsTree.SelectedNode = ComponentsTree.GetNodeAt(e.X, e.Y);
                if (ComponentsTree.SelectedNode != null)
                {
                    contextMenuStrip1.Show(ComponentsTree, e.Location);
                }
            }
        }

        private bool checkForRecursion(TreeNode parentNode, String currNodeName)
        {
            if (parentNode != null)
                if (returnNameOfTheNode(parentNode) != currNodeName)
                    return checkForRecursion(parentNode.Parent, currNodeName);
                else
                    return false;
            else
                return true;
        }

        private bool isContainsInDB(String componentName)
        {
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
            {
                bool maintainsFlag = false;
                connection.Open();
                String inComponents = "SELECT * FROM INCOMP WHERE NAME = '" + componentName + "'";
                SqlCommand sqlCommand = new SqlCommand(inComponents, connection);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                if (reader.HasRows)
                    maintainsFlag = true;
                reader.Close();

                String topComponents = "SELECT * FROM TOPCOMP WHERE NAME = '" + componentName + "'";
                SqlCommand sqlCommand1 = new SqlCommand(topComponents, connection);
                reader = sqlCommand1.ExecuteReader();
                if (reader.HasRows)
                    maintainsFlag = true;
                reader.Close();

                connection.Close();
                return maintainsFlag;
            }
        }

        private String returnNameOfTheNode(TreeNode node)
        {
            String nonRegEx = node.Text;
            string pattern = @" \(\d*\)";
            string target = "";
            Regex regex = new Regex(pattern);
            return regex.Replace(nonRegEx, target);
        }

        private int returnQuantityOfTheNode(TreeNode node)
        {
            String nonRegEx = node.Text;
            string pattern = @"\(\d*\)";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(nonRegEx);
            String result;
            if (match.Success)
            {
                result = match.Value;
                return int.Parse(result.Substring(1, result.Length - 2));
            }
            else return 0;                
        }

        private int getParentId(TreeNode node)
        {
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
            {
                connection.Open();
                int parentId = 0;
                String parentName = returnNameOfTheNode(node.Parent);
                String command = "SELECT ID FROM TOPCOMP WHERE NAME = '" + parentName + "'";
                SqlCommand sqlCommand = new SqlCommand(command, connection);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    parentId = (int)reader.GetValue(0);
                }
                reader.Close();

                command = "SELECT ID FROM INCOMP WHERE NAME = '" + parentName + "'";
                sqlCommand = new SqlCommand(command, connection);
                reader = sqlCommand.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    parentId = (int)reader.GetValue(0);
                }
                reader.Close();

                connection.Close();
                return parentId;
            }
        }

        private int getNodeId(TreeNode node)
        {
            String componentName = returnNameOfTheNode(node);
            return getNodeIdByName(componentName);
        }

        private int getNodeIdByName(String name)
        {
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
            {
                connection.Open();
                String command = "SELECT ID FROM INCOMP WHERE NAME = '" + name + "'";
                SqlCommand sqlCommand = new SqlCommand(command, connection);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    int Id = (int)reader.GetValue(0);
                    reader.Close();
                    connection.Close();
                    return Id;
                }
                reader.Close();

                command = "SELECT ID FROM TOPCOMP WHERE NAME = '" + name + "'";
                sqlCommand = new SqlCommand(command, connection);
                reader = sqlCommand.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    int Id = (int)reader.GetValue(0);
                    reader.Close();
                    connection.Close();
                    return Id;
                }
                reader.Close();
                connection.Close();

                return 0;
            }
        }

        private TreeNode[] findNodeByName(TreeNodeCollection nodes, String name)
        {
            TreeNode[] returnnodes = nodes.Cast<TreeNode>().Where(r => returnNameOfTheNode(r) == name).ToArray();
            return returnnodes;
        }

    }
}

