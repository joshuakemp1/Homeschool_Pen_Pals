# Start from the first row
# Find the 'Matched To' Column on that row
# Search the next rows until you find the same name as the 'Name' field
# Grab the info from the row ( matchchild row ) and send it in an email
# Go back to the next row in the excel and search for the next 'Matched To'
require 'roo'
require 'pry-byebug'
require 'gmail'

# Take 3 inputs from the command line:
# Email
# Password
# Sheet name
puts "Enter the email you are sending the emails from ( probably gratefulwithtwo@gmail.com )"
script_email = gets.chomp
puts "Enter the email's password: "
script_email_password = gets.chomp
puts "Enter the sheet name that we want to run this script against NOTE: be VERY sure you get the correct name or else you will RESEND old emails WATCH SPACING!"
sheet_name = gets.chomp

puts "This is the script email: #{script_email}, this is the script email password: #{script_email_password}, this is the sheet name: #{sheet_name}"

# Have a loop that grabs the first matched child's name, email, and match to(buddy) and store those as variables- this is who we are sending the email to
# Now inside that loop search the name column and find the match to(buddy) - grab the match to (buddy info to send in email )

@spreadsheet = Roo::Spreadsheet.open("C:/Users/joshua.kemp/Documents/UI path/HomeSchoolPenPals/Homeschool Pen Pals Latest.xlsx", extension: :xlsx)
@spreadsheet.default_sheet = sheet_name # switch to the sheet name that was passed as an ARG

for row in @spreadsheet.parse
    original_child_name = row[1].to_s
    matched_childs_name = row[9].to_s
    child_email = row[5].to_s

    for row_2 in @spreadsheet.parse
     child_name = row_2[1].to_s # this loop just keeps going down the MATCHED TO column

        if child_name == matched_childs_name
            puts "We have a match: Child name: #{child_name} and #{matched_childs_name}"  

            #  Grab all of the matched to child's data
            matched_childs_age = row_2[2].to_s
            matched_childs_gender = row_2[3].to_s
            matched_childs_address = row_2[4].to_s
            matched_childs_parents_email = row_2[5].to_s
            matched_childs_parents_IG_handle = row_2[6].to_s
            matched_childs_interests = row_2[7].to_s
            matched_childs_other_interests = row_2[8].to_s

            email_body = "This is all of the matched child's info: name = #{original_child_name}, gender = #{matched_childs_gender}, age = #{matched_childs_age}, parents email = #{matched_childs_parents_email}, parents IG handle = #{matched_childs_parents_IG_handle}, Address = #{matched_childs_address}, Interests 1 = #{matched_childs_interests}, Interests 2 = #{matched_childs_other_interests} "

            puts "SENT #{original_child_name} email! Here's the email address it was sent to: #{child_email}"
            # Send Email
            gmail = Gmail.connect(script_email, script_email_password)

            email = gmail.compose do
            to child_email # Change this to child_email when we are ready to test for real
            subject "Pen Pals Match!"
                html_part do
                    content_type 'text/html; charset=UTF-8'
                    body "<span><br>Hi, Welcome To HomeSchool Pen Pals, <b>"+original_child_name+"</b> has been matched!</br></span><span><br><br>Here is your Pen Pal's information: </br></br></span><span><br><br><b>Name: "+matched_childs_name+"</b></br></br></span><span><b><br><br>Gender: "+matched_childs_gender+"</b></br></br></span><span><b><br><br>Age: "+matched_childs_age+"</b></br></br></span><span><b><br><br>Parent's Email: "+matched_childs_parents_email+"</b></br></br></span><span><b><br><br>Parent's IG Handle: "+matched_childs_parents_IG_handle+"</b></br></br></span><span><b><br>Address: </b>"+matched_childs_address+"</br></span><span><b><br>Interests: </b>"+matched_childs_interests+";"+matched_childs_other_interests+"</br></span><span><br><br>Thanks so much for participating and be sure to use the hashtag #homeschoolpenpals on IG so we can all enjoy this beautiful community together!</br></br></span><span><br>xo, Elisha</br></span><span><br><a href=http://www.instagram.com/elishakemp>Personal IG</a>   <b>|</b>   <a href=http://bit.ly/wildflowernest>Wildflower Nest</a></br></span>"
                end
            end
            email.deliver!
            gmail.logout
        else

        end
    end
end