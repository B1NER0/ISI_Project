﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using Newtonsoft.Json;
using RestSharp;
using RestSharp.Authenticators;
using System.Diagnostics;
using Newtonsoft;
using Newtonsoft.Json.Linq;

namespace ProjectManagementToolkit
{
    public partial class SignUp : Form
    {
        public string DBConnectionString;
        private const string BASE_URL = "https://kanban-api-624.herokuapp.com/";
        public SignUp()
        {
            InitializeComponent();
            DBConnectionString = Properties.Settings.Default.ISI_DBConnectionString;
        }

        private void chkShowPassword_CheckedChanged(object sender, EventArgs e)
        {
            //Toggle confirmation password's visibility
            if (chkShowPassword.Checked)
                txtConfirmPassword.PasswordChar = '\0';
            else
                txtConfirmPassword.PasswordChar = '*';
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

        }

        private void btnSignUp_Click(object sender, EventArgs e)
        {
            //Store input in variabless
            string userName = txtUsername.Text,
                   password = txtPassword.Text,
                   confirmPassword = txtConfirmPassword.Text;

            //Hide error labels
            lblUsernameError.Visible = false;
            lblPasswordError.Visible = false;

            //Check for username errors
            if (userName == "")
            {
                lblUsernameError.Text = "Username field cannot be empty";
                lblUsernameError.Visible = true;
                return;
            }
            else if (userName.Length > 50) //Ensure that username length is below cap
            {
                lblUsernameError.Text = "Username field cannot be longer than 50 characters";
                lblUsernameError.Visible = true;
                return;
            }

            //Check password for errors
            if (password == "")
            {
                lblPasswordError.Text = "Password field cannot be empty";
                lblPasswordError.Visible = true;
                return;
            }

            //Check if password and confirm password fields match
            if (!(password == confirmPassword))
            {
                lblPasswordError.Text = "Password field and Confirm Password field do not match";
                lblPasswordError.Visible = true;
                return;
            }

            if (Validation.CheckIfUserExists(userName))
            {
                lblUsernameError.Text = "This username already exists";
                lblUsernameError.Visible = true;
                return;
            }

            if(create_user(userName,password,"admin") && create_local_user(userName,password))
            {
                MessageBox.Show("You have successfully created an account!");
            }
            else
            {
                MessageBox.Show("Account creation failed!");
            }
            
        }

        private void SignUp_Load(object sender, EventArgs e)
        {
            txtUsername.Text = "Username";
            txtPassword.Text = "Password";
            txtConfirmPassword.Text = "Password";
            txtUsername.BackColor = this.BackColor;
            txtPassword.BackColor = this.BackColor;
            txtConfirmPassword.BackColor = this.BackColor;
        }

        public bool create_user(string username, string password, string role)
        {
            try
            {
                var client = new RestClient(BASE_URL);
                var request = new RestRequest("/user/signup", Method.POST);
                request.RequestFormat = DataFormat.Json;
                request.AddHeader("Content-type", "application/json");
                request.AddJsonBody(new
                {
                    username = username,
                    password = password,
                    role = role
                });

                var response = client.Execute(request);
                HttpStatusCode statusCode = response.StatusCode;
                Debug.WriteLine("login_user" + response.Content);
                int num_status_code = (int)statusCode;
                if (num_status_code == 201)
                {
                    return true;
                }


            }
            catch (Exception e)
            {
                Debug.WriteLine(e.ToString());
            }
            return false;
        }

        private bool create_local_user(string userName, string password)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnectionString))
                {
                    //If no user with the same username exists, add the user to the database
                    string insertUser = "INSERT INTO users (username, hashedpassword) VALUES (@userName, @hashedPassword)";
                    using (SqlCommand insertUserCommand = new SqlCommand(insertUser))
                    {
                        insertUserCommand.Connection = conn;
                        insertUserCommand.Parameters.Add("@userName", SqlDbType.VarChar, 50).Value = userName;
                        insertUserCommand.Parameters.Add("@hashedPassword", SqlDbType.NChar, 20).Value = Hashing.HashPassword(password);

                        conn.Open();
                        insertUserCommand.ExecuteNonQuery();

                        conn.Close();                       
                        Close();
                        return true;
                    }
                }
            }
            catch(SqlException e)
            {
                MessageBox.Show(e.Message);
            }
            return false;
        }
    }
}
