import os

uuids = ["67ab0686d32c11eb91880a4165e7dae6",
         #"138ff1e2d26e11eb91880a4165e7dae6",
         #"d5475046d4b711eb91880a4165e7dae6",
         "c8387768d58911eb91880a4165e7dae6",
         "6a5a8352d7e111eb91880a4165e7dae6",
         "d4ac7dd4d8b611eb91880a4165e7dae6",
         "3964aa42f90c11eb91880a4165e7dae6",
         #"9b390dce600f11ed8b27b13086acc439",
         #"f7ac87d060c711ed8b27b13086acc439",
         #"13768e14640d11ed8b27b13086acc439",
         #"86a472c464c711ed8b27b13086acc439",
         #"a023b58c65a511ed8b27b13086acc439",
         #"7bba6ec865bb11ed8b27b13086acc439"
         ]


for uuid in uuids:
    os.system('python pr_db_gen.py --uuid {0} --date_start {1} --date_end {2}'.format(uuid,'2023-04-01','2023-04-30'))




    
