# Automate-prutor.com
This is a pre-WATSON script written in python to automate the registering process on prutor.com

Important points to make sure the script runs :-

If the script donot runs :-

* The script makes use of computer vision and is specifically programmed keeping the developer's screen resolution so it might not work on other devices. If that is the case WATSON Hyperpersonalised Tool can be used to build own script from scratch (contact developer.)

If the script runs :-

* feed the data in student.xlsx as per schema. (an example file is provided in the attachment. Remember to keep the schema as per the example file.)
* Arrange all the roll_numbers in assending order (Add missing roll_numbers and put their name as null).
* the roll_number is assumed to be alphanumeric so enter the fixed alpha part in the initial parameter and last expected numerical value as first parameter.  i.e; if roll_numers are like 19J010, 19J011,......,19J099. then the fixed part is 19J0 and the changing part is 10,11,......,99. So enter  99 as the first parameter so the program runs until it has fetched the 19J099  roll_number.
* Set the "initialschno" value to the (starting roll_number - 1) i.e; if roll (the numeric part of the roll_number ) starts from 3 then put 2 in "initialschno".
* Set the section for which uploading is being done.

For the example file provided :-

* Values for 4 fields are already provided.
* The rolls are in assending order.
* initial part is 19J and first parameter goes as 102 (Last roll_number being 19J102)
* "initialschno"  is set as 99 (100-1 as the first roll_number is 19J100)
* section is J (as guessed by seeing the roll_number).

For any further info reach developer at shashanksharma191098@gmail.com
