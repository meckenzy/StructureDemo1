Public Class Form1

    'form level members
    Private objCustomers As New ArrayList
    
    Public Sub CreateCustomer(firstName As String, lastName As String, email As String)
        'Declare a Customer object
        Dim objNewCustomer As Customer

        'Create the new customer
        objNewCustomer.FirstName = firstName
        objNewCustomer.LastName = lastName
        objNewCustomer.Email = email

        'Add the new customer to the list
        objCustomers.Add(objNewCustomer)

        'Add the new customer to the ListBox control 
        lstCustomers.Items.Add(objNewCustomer)

    End Sub
    Public Sub DisplayCustomer(customer As Customer)
        'Display the customer details on the form
        txtName.Text = customer.Name
        txtFirstName.Text = customer.FirstName
        txtLastName.Text = customer.LastName
        txtEmail.Text = customer.Email
    End Sub
    Private Sub btnListCustomer_Click(sender As Object, e As EventArgs) Handles btnListCustomer.Click
        'Create a new customer
        Dim objCustomer As Customer
        'objCustomer.FirstName = "Serge"
        'objCustomer.LastName = "Badio"
        'objCustomer.Email = "sbadio@hotmail.com"

        'Create some customers
        CreateCustomer("Serge", "Badio", "sbadio@somecompany.com")
        CreateCustomer("Frank", "Peoples", "fpeoples@somecompany.com")
        CreateCustomer("Bill", "Scott", "bscott@somecompany.com")

        'Display the customer
        DisplayCustomer(objCustomer)
    End Sub
    Public ReadOnly Property SelectedCustomer() As Customer
        Get
            If lstCustomers.SelectedIndex <> -1 Then
                'Return the slected customer
                Return CType(objCustomers(lstCustomers.SelectedIndex), Customer)
            End If
        End Get
    End Property

    Private Sub btnDeleteCustomer_Click(sender As Object, e As EventArgs) Handles btnDeleteCustomer.Click
        'If no customer is selected in the ListBox then
        If lstCustomers.SelectedIndex = -1 Then
            'Display a message
            MessageBox.Show("You must select a customer to delete.", "Structure Demo")

            'Exit this method
            Exit Sub
        End If

        'Promt the user to delete the selected customer
        If MessageBox.Show("Are you sure you want to delete " &
            SelectedCustomer.Name & "?", "Structure Demo",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) =
            DialogResult.Yes Then

            'Get the customer to be deleted
            Dim objCustomerToDelete As Customer = SelectedCustomer

            'Remove the customer from the ArrayList
            objCustomers.Remove(objCustomerToDelete)

            'Remove the customer from the ListBox
            lstCustomers.Items.Remove(objCustomerToDelete)

        End If
    End Sub

    Private Sub lstCustomers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCustomers.SelectedIndexChanged
        'Display the customer details
        DisplayCustomer(SelectedCustomer)
    End Sub
End Class
