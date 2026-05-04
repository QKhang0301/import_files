Problem:
    I the final phase of my job, I have to import images and some types of files into Excel.
    However, via insert option in excel panel, it only allow insert one object at a time
Solution:
    -> Write a script that automatically import files, images into excel, which is a pre-defined template
    -> I only write a command line with arguments, then it will create a excel template, then import images or files into it
Define a problem:
    There are 4 types of templates
        1. BAT template -> need to import and pre-defined images (Images that has specific names)
        2. CP_PDC template -> need images, .zip, .xlsx, .html, .log, and .msg file
        3. CONT_CHECKLIST template -> need imagee, .zip and .xlsx file
        4. CONT_PDC template -> need images , .html, .log, and .msg file

    Each imported file has its own name, the name contains specific keyword
-------------------------------------------------------------------------------------
    For BAT, it contains images:
        itm, poc, algoID
        sqm_open, sqm_bat, sqm_gnd, sqm_cross
        swm_open, swm_bat, swm_gnd, swm_cross
        psem_open, psem_bat, psem_gnd, psem_cross
        aod_gnd
        firing_result, firing_value

        Besides, It contains other information:
            Snapshot name
            Project name
            ECU type
            Harness type
            Hardware and software phase
            Sequence ID
            BB number
            SW number
            SW version
-------------------------------------------------------------------------------------



