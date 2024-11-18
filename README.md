=IF(AND(C2="Yes", D2="Yes"), "Complete and Correct Information", IF(AND(C2="Yes", D2="No"), E2, IF(AND(C2="No", D2="Yes"), "Incomplete Info", IF(AND(C2="No", D2="No"), E2, ""))))
