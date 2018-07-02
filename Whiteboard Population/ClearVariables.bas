Attribute VB_Name = "ClearVariables"
Option Private Module
Public Sub clear_vars(job_number_array, serial_number, year_job_number, job_number, customer_name, model_number, _
                        budget_file_path, budget_file_name, file_path, number_of_jobs, cab_hours, electrical_hours, _
                        refrigeration_hours, job_type)

    'Clear variables
    Erase job_number_array
    serial_number = ""
    year_job_number = ""
    job_number = ""
    customer_name = ""
    model_number = ""
    budget_file_path = ""
    budget_file_name = ""
    file_path = ""
    job_type = ""
    number_of_jobs = 0
        job_counter = 0
    cab_hours = 0
    electrical_hours = 0
    refrigeration_hours = 0

End Sub
