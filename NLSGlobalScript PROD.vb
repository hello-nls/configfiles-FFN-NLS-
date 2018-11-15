'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GLOBAL CONSTANTS                                                                                                     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' * Loan 
'		Portfolios
			Public Const PORT_CODE_NUM_CPLUS = 1			
			Public Const PORT_CODE_NUM_FPLUS = 0
' 		Status Codes
			Public Const LOAN_STATUS_ACTIVE = 0
			Public Const LOAN_STATUS_CLOSED = 1
			Public Const LOAN_STATUS_RESTRUCTURE = 2
			Public Const LOAN_STATUS_DRAFT = 3
			Public Const LOAN_STATUS_CANCELLED = 4
			Public Const LOAN_STATUS_GL_NON_ACCRUAL = 5
			Public Const LOAN_STATUS_GL_CHARGE_OFF = 6
			Public Const LOAN_STATUS_CURRENT = 10 
			Public Const LOAN_STATUS_CHARGED_OFF = 12 
			Public Const LOAN_STATUS_COMPLETED = 13 
			Public Const LOAN_STATUS_NEAR_PAYOFF = 14 
			Public Const LOAN_STATUS_DELETE = 15 
			Public Const LOAN_STATUS_BK_REPORTED = 16 
			Public Const LOAN_STATUS_DECEASED_REPORTED = 17 
			Public Const LOAN_STATUS_COLLECTIONS = 18 
			Public Const LOAN_STATUS_REFUSAL_TO_PAY = 19 
			Public Const LOAN_STATUS_INTENT_TO_PAY = 20 
			Public Const LOAN_STATUS_MILITARY = 21 
			Public Const LOAN_STATUS_INCARCERATION = 22 
			Public Const LOAN_STATUS_DEFAULTED_120_C = 23 
			Public Const LOAN_STATUS_DEFAULTED_120_F = 24 
			Public Const LOAN_STATUS_DEFAULTED_BK_CONFIRMED = 25 
			Public Const LOAN_STATUS_DECEASED = 26 
			Public Const LOAN_STATUS_FRAUD = 27 
			Public Const LOAN_STATUS_FIRST_PAY_DEFAULT = 28 
			Public Const LOAN_STATUS_P_STATEMENTS = 29 
			Public Const LOAN_STATUS_PENDING_COMPLETE = 30 
			Public Const LOAN_STATUS_APPROVED_SETTLEMENT = 31 
			Public Const LOAN_STATUS_ACTIVE_SETTLEMENT = 32 
			Public Const LOAN_STATUS_COMPLETED_SETTLEMENT = 33 
			Public Const LOAN_STATUS_BROKEN_SETTLEMENT = 34 
			Public Const LOAN_STATUS_PENDING_COMPLETE_DS = 35 
			Public Const LOAN_STATUS_MODIFICATION_APPROVED = 36 
			Public Const LOAN_STATUS_CEASE_AND_DESIST_ACTIVE = 37 
			Public Const LOAN_STATUS_INACTIVE_POA = 38 
			Public Const LOAN_STATUS_POA_RECEIVED = 39 
			'Public Const LOAN_STATUS_SOLD = 40 
			Public Const LOAN_STATUS_BK_NOTICE = 41 
			Public Const LOAN_STATUS_BK_FILED = 42 
			Public Const LOAN_STATUS_BK_ACTIVE = 42
			Public Const LOAN_STATUS_BK_DISMISSED = 43 
			Public Const LOAN_STATUS_BK_COMPLETED = 44
			Public Const LOAN_STATUS_BK_DISCHARGE7 = 46 
			Public Const LOAN_STATUS_BK_DISCHARGE11 = 47  
			Public Const LOAN_STATUS_BK_DISCHARGE12 = 48  
			Public Const LOAN_STATUS_BK_DISCHARGE13 = 49 
			Public Const LOAN_STATUS_APPROVAL_REQUIRED = 10000
			Public Const LOAN_STATUS_APPROVAL_REQUESTED = 10001
' 		Transaction Codes
			Public Const LOAN_TRANS_CODE_LATE_FEE = "150"
			Public Const LOAN_TRANS_CODE_NSF_FEE = "180"
			Public Const LOAN_TRANS_CODE_CHECK_FEE = "504"
			Public Const LOAN_TRANS_CODE_LATE_FEE_PAYMENT = "250"
			Public Const LOAN_TRANS_CODE_NSF_FEE_PAYMENT = "280"
			Public Const LOAN_TRANS_CODE_CHECK_FEE_PAYMENT = "506"	
			
' * Contact
'		Types
'		CONTACT DETAILS
			Public Const CIF_DETAIL_GENERAL_ADMIN_PAGE_CIFNO = 164373

' * Task
'		Templates
			Public Const TEMPLATE_NUM_LOAN_MODIFICATIONS = 1           
			Public Const TEMPLATE_NUM_LOAN_PAYOFF = 2                  
			Public Const TEMPLATE_NUM_DELETE_LOAN = 3                  
			Public Const TEMPLATE_NUM_SELECT_A_TASK = 4                
			Public Const TEMPLATE_NUM_PAYMENT_CHANGE = 5               
			Public Const TEMPLATE_NUM_DEBT_SETTLEMENT = 6              
			Public Const TEMPLATE_NUM_LOAN_MODIFICATIONS_II = 7        
			Public Const TEMPLATE_NUM_BANKRUPTCY = 8                    
			Public Const TEMPLATE_NUM_APPROVE_NEW_LOAN = 50000         
			Public Const TEMPLATE_NUM_APPROVE_LOAN_MODIFICATION = 50001
			Public Const TEMPLATE_NUM_APPROVE_TRANSACTION = 50002      
'		Statuses                                                       
			Public Const TEMPLATE_STATUS_CREATE_REQUEST = 1            
			Public Const TEMPLATE_STATUS_SUBMIT_FOR_APPROVAL = 2       
			Public Const TEMPLATE_STATUS_PAUSED = 3                    
			Public Const TEMPLATE_STATUS_DENIEDB = 4                    
			Public Const TEMPLATE_STATUS_DENIED_WCOUNTER = 5           
			Public Const TEMPLATE_STATUS_APPROVAL_PENDING = 6          
			Public Const TEMPLATE_STATUS_STARTED = 7                   
			Public Const TEMPLATE_STATUS_COMPLETED = 8                 
			Public Const TEMPLATE_STATUS_CREATED_NO_DOC = 9            
			Public Const TEMPLATE_STATUS_DELETION_COMPLETE = 10        
			Public Const TEMPLATE_STATUS_SUBMITTED_FOR_APPROVAL = 11   
			Public Const TEMPLATE_STATUS_ACTIVE_SETTLEMENt = 12        
			Public Const TEMPLATE_STATUS_PENDING_DOCS = 13  
			Public Const TEMPLATE_STATUS_BK_NOTICE = 17
			Public Const TEMPLATE_STATUS_BK_DISSMISSED = 15	
			Public Const TEMPLATE_STATUS_BK_ACTIVE = 16 'later replaced TEMPLATE_STATUS_BK_FILED	
			Public Const TEMPLATE_STATUS_BK_FILED = 16
			Public Const TEMPLATE_STATUS_BK_COMPLETED = 14		
			Public Const TEMPLATE_STATUS_NOT_STARTED = 100001          
			Public Const TEMPLATE_STATUS_APPROVED = 100002             
			Public Const TEMPLATE_STATUS_CANCELLED = 100003            
			Public Const TEMPLATE_STATUS_DENIED = 100004               
'		Queues                                                                                    
			Public Const TEMPLATE_QUEUE_REQUESTS = 1  
			Public Const TEMPLATE_QUEUE_APPROVALS = 2
			Public Const TEMPLATE_QUEUE_LOAN_DELETES = 3
			Public Const TEMPLATE_QUEUE_SETTLEMENTS = 4
			Public Const TEMPLATE_QUEUE_BANKRUPTCY = 5
			Public Const TEMPLATE_QUEUE_MODIFICATIONS = 6
			
' * Document Templates
			Public Const TEMPLATE_DOCTEMPL_FPLUS_DEBT_VALIDATION_LETTER = 11 
			Public Const TEMPLATE_DOCTEMPL_CPLUS_DEBT_VALIDATION_LETTER = 10
			
'Other Constants
    'Security
			Public Const SECURITY_GROUP_MODIFICATION_APPROVERS = 22
			Public Const SECURITY_GROUP_ACH_FILE_IMPORT = 23
			Public Const SECURITY_GROUP_COLLECTIONS_LEADERS = "(5,8,20)"
                                                Public Const SECURITY_GROUP_SERVICING_LEADERS = "(2,4,20)"
			Public Const SECURITY_GROUP_COLLECTION_OWNERS = 24
    'Folders and Directories
			Public Const INVESTOR_TRANSFERS_PENDING_DIR = "\\tsclient\T\ACH\Investor Transfers\PendingTransfers\"
			Public Const INVESTOR_TRANSFERS_COMPLETED_DIR = "\\tsclient\T\ACH\Investor Transfers\CompletedTransfers\"
			Public Const EXPERIAN_IMPORTS_PENDING_DIR = "\\tsclient\T\FreedomPlus Servicing\NLS file imports\Experian Source files\Pending"
			Public Const EXPERIAN_SCORE_IMPORTS_COMPLETED_DIR = "\\tsclient\T\FreedomPlus Servicing\NLS file imports\Completed Imports\Score Imports\"
			Public Const EXPERIAN_SKIP_IMPORTS_COMPLETED_DIR = "\\tsclient\T\FreedomPlus Servicing\NLS file imports\Completed Imports\Skip Imports\"
  

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GLOBAL Scripting VARIABLES                                                                                                          
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
PENDING_COMPLETE_CLOSE_LOAN_SQL = "Select 'closuredetails' =  l.loan_number + ',' + convert(varchar(15), l.acctrefno) + ',' +                          " & chr(13) & _
                      "Case When l.current_suspense > 0 Then cast(convert(numeric(10,2),round(l.current_suspense,2)) as varchar)   Else '0' End + ',' +" & chr(13) & _
                      "CASE When l.portfolio_code_id = 0 then '5'  When l.portfolio_code_id = 1 then '6' End + ',' +                                   " & chr(13) & _
                      "CASE When                                                                                                                       " & chr(13) & _
                      "   (CASE when l.portfolio_code_id = 0 then ld.userdef18                                                                         " & chr(13) & _
                      "         When l.portfolio_code_id = 1 then ld.userdef17                                                                         " & chr(13) & _
                      "		 else '0'   End) = '1' or c.email is null Then 'TRUE'                                                                      " & chr(13) & _
                      "    Else 'FALSE' End + ',' +                                                                                                    " & chr(13) & _
                      "CASE When c.email is null Then '' Else c.email end + ',' +" & chr(13) & _	
                      "c.firstname1+' '+lastname1 + ',' +l.loan_number + ',' +c.firstname1 + ',' +lastname1                           " & chr(13) & _
                      " INTO #loanstoclose                                                                                                             " & chr(13) & _
                      "from                                                                                                                            " & chr(13) & _
                      "   cif c, loanacct_detail ld,                                                                                                   " & chr(13) & _
                      "   loanacct_statuses ls,                                                                                                        " & chr(13) & _
                      "   loanacct l                                                                                                                   " & chr(13) & _
                      "                                                                                                                                " & chr(13) & _
                      "where                                                                                                                           " & chr(13) & _
                      "   l.acctrefno = ls.acctrefno                                                                                                   " & chr(13) & _
                      "   and l.acctrefno = ld.acctrefno                                                                                               " & chr(13) & _
                      "   and c.cifno = l.cifno                                                                                                        " & chr(13) & _
                      "   and l.status_code_no = 0                                                                                                     " & chr(13) & _
                      "   and ls.status_code_no = 30 /*PENDING COMPLETE status code no */                                                              " & chr(13) & _
                      "   and l.current_payoff_balance <= 0                                                                                            " & chr(13) & _
                      "   and l.current_pending <= 0                                                                                                   " & chr(13) & _
                      "   and convert(char(8),ls.effective_date,1) <= dateadd(dd,-7,getdate())                                                         " & chr(13) & _
                      "                                                                    " & chr(13) & _
                      "SELECT (SELECT count(*) - 1 From #loanstoclose) as loancnt, closuredetails from #loanstoclose   order by closuredetails         " & chr(13) & _
                      "Drop table #loanstoclose " & chr(13) & _
                      "--This SQL looks for any laons with a PENDING COMPLETE status for a selected number of days(currently 14) and generates an XML in NLS format that will removed the" & chr(13) & _ 
                      "--PENDING COMPLETE status and add the COMPLETE status and close the loan. "

'ARRAYS
Dim investormatrix(40)
'key - [class2 codenum],[class2 name],[F+ ach company ID],[F+ ach company name],[C+ ach company ID],[C+ ach company name]

	investormatrix(0) = "0,,,,,"
	investormatrix(1) = "1,NON-ACH,,,,"
	investormatrix(2) = "2,1 - VULCAN,1,Vulcan F+,4,Vulcan C+"
	investormatrix(3) = "3,2 - AEQUITAS,,,,"
	investormatrix(4) = "4,3 - FFNAM FUND,3,FFAM,3,FFAM"
	investormatrix(5) = "5,2A - AEQUITAS A,,,,"
	investormatrix(6) = "6,2B - AEQUITAS B,,,,"
	investormatrix(7) = "7,4 - RANGER,7,Ranger F+,7,Ranger F+"
	investormatrix(8) = "8,5 - ATALAYA CAPITAL,8,Atalaya,8,Atalaya"
	investormatrix(9) = "9,6 - RANGER DLT,9,Ranger DLT F+,9,Ranger DLT F+"
	investormatrix(10) = "10,7 - BLUE ELEPHANT,10,Blue Elephant F+,10,Blue Elephant F+"
	investormatrix(11) = "11,8 - TIDEPOOL,11,TidePool F+,11,TidePool F+"
	investormatrix(12) = "12,9 - AEQUITAS ACC5,12,Aequitas ACC5,15,AEQUITAS ACC5 C+"
	investormatrix(13) = "13,5A - ATALAYA V-B,13,Atalaya V-B,13,Atalaya V-B"
	investormatrix(14) = "14,5B - ATALAYA VI,,,,"
	investormatrix(15) = "15,5C - ATALAYA MW,16,Atalaya MW,16,Atalaya MW"
	investormatrix(16) = "16,7A - BLUE ELEPHANT OFFSHO,17,Blue Elephant Offshore F+,17,Blue Elephant Offshore F+"
	investormatrix(17) = "17,5D - ATALAYA II-C,18,Atalaya II-C,18,Atalaya II-C"
	investormatrix(18) = "18,10 - KEYSTONE,19,KEYSTONE,19,KEYSTONE"
	investormatrix(19) = "19,11 - CRESTLINE,20,CRESTLINE,20,CRESTLINE"
	investormatrix(20) = "20,12 - FCCF-SERIES 1,21,FCCF-Series 1,21,FCCF-Series 1"
	investormatrix(21) = "21,13 - FCCF-SERIES 2,22,FCCF-Series 2,22,FCCF-Series 2"
	investormatrix(22) = "22,14 - ACM III,23,ACM III,23,ACM III"
	investormatrix(23) = "23,15 - ACM MW II,24,ACM MW II,24,ACM MW II"
	investormatrix(24) = "24,16 - CRESTLINE II,25,CRESTLINE II,25,CRESTLINE II"
	investormatrix(25) = "25,17 - DAKOY,26,DaKoy Income Fund LP,26,DaKoy Income Fund LP"
	investormatrix(26) = "26,18 - STONERIDGE,27,StoneRidge,27,StoneRidge"
	investormatrix(27) = "27,19 - ACM ALAMOSA I,32,ACM Alamosa I,32,ACM Alamosa I"
	investormatrix(28) = "28,20 - ACM ALAMOSA I-A,28,ACM ALAMOSA I-A,28,ACM ALAMOSA I-A"
	investormatrix(29) = "29,21 - ACM GATEWAY,30,ACM Gateway,30,ACM Gateway"
	investormatrix(30) = "30,22 - ACM F ACQUISITION,31,ACM F Acquisition,31,ACM F Acquisition"
	investormatrix(31) = "31,23 - 590 CONSUMER LENDING,29,590 Consumer Lending Corp,29,590 Consumer Lending Corp"
	investormatrix(32) = "32,18A - STONE RIDGE II,34,Stone Ridge II F+,34,Stone Ridge II F+"
	investormatrix(33) = "33,25 - ACM IV OFF,35,ACM IV Off F+,38,ACM IV Off C+"
	investormatrix(34) = "34,24 - ACM IV ON,36,ACM IV On F+,37,ACM IV On C+"
	investormatrix(35) = "35,13A - FCCF - CS,39,FCCF-CS,39,FCCF-CS"
	investormatrix(36) = "36,13B-FCCF-ABS1,41,FCCF-ABS1 F+,40,FCCF-ABS1 C+"
	investormatrix(37) = "37,13Z-FCCF-INEL,43,FCCF-INEL F+,42,FCCF-INEL C+"
	investormatrix(38) = "38,26-STONERIDGE- ALITI,44,StoneRidge - ALITI,44,StoneRidge - ALITI"
	investormatrix(39) = "39,13C-FCCF-ABS2,45,FCCF-ABS2,45,FCCF-ABS2" 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GLOBAL SUBS                                                                                                          
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FirstPayDefaultFlag()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'First Pay Default Flag UDF Maintainance
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    timecheck = NLSapp.SQLSelectStatement("SELECT Case When substring(cd.userdef03,10,25) > getdate() Then 'N' Else 'Y' End FROM cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'FIRST PAY DEFAULTS'")
    
    If timecheck = "Y" Then
        xmlsql = NLSapp.SQLSelectStatement("SELECT cd.userdef18 FROM cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'FIRST PAY DEFAULTS'")
        
        detail2udf13xml = NLSapp.SQLSelectStatement(xmlsql)
        
        If detail2udf13xml = "" Then
            detail2udf13xml = "***END OF RESULT***"
        End If
        
        Do Until detail2udf13xml = "***END OF RESULT***"
        
           ImportXML(detail2udf13xml)
            'Expected XML string output Format:
            '<NLS EnforceTagExistence="1">
            '  <LOAN  UpdateFlag="1" AcctRefno="999999" >
            '    <LOANDETAIL2 UserDefined13 = "1" />
            '  </LOAN>
            '</NLS>
        
            detail2udf13xml = NLSApp.SQLSelectStatementReadNext()
        
        Loop
        'Update last/Next Run datetime and Last user record
        Runstamp = NOW()
        NextRun = Date() + 1 & " 8:00:00 AM"
         NextRun = NOW()
        UID = NLSApp.GetCurrentUserNo()
        
        Cifupdate = "<NLS><CIF UpdateFlag='1'CIFNumber='FIRST PAY DEFAULTS'><CIFDETAIL CIFDetail3='Next Run: " & NextRun & _
            "' CIFDetail4='Last Run: " & Runstamp & "' CIFDetail5=' UserID: " & UID & "' /></CIF></NLS>"
        
        ImportXML(Cifupdate)
    End if

End Sub
'-------------------------------------------------------------------------------------------------------------------------

Sub AutoWOaddCompletePending()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Automatic Write-off loan balances $10 or less set loan status to COMPLETE PENDING
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    xmlsql = NLSapp.SQLSelectStatement("SELECT cd.userdef18 FROM cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'AUTO WRITEOFFCLOSE'")

    writeoffxml = NLSapp.SQLSelectStatement(xmlsql)

    If len(writeoffxml) = 0 Then
        writeoffxml = "***END OF RESULT***"
    End If

    Do Until writeoffxml = "***END OF RESULT***"

        ImportXML(writeoffxml)
        'Expected XML string output Format:
        '<NLS>                                                                                                         
        '  <TRANSACTIONS>                                                                                              
        '    <TRANSACTIONCODE TransactionCode="540" EffectiveDate="07/05/17" LoanNumber="APP-00099999" Amount="8.84" />
        '    <TRANSACTIONCODE TransactionCode="544" EffectiveDate="07/05/17" LoanNumber="APP-00099999" Amount="8.84" />
        '    <TRANSACTIONCODE TransactionCode="448" EffectiveDate="07/05/17" LoanNumber="APP-00099999" Amount="8.84" />
        '    <TRANSACTIONCODE TransactionCode="452" EffectiveDate="07/05/17" LoanNumber="APP-00099999" Amount="8.84" />                        
        '  </TRANSACTIONS>                                                                                             
        '  <LOAN UpdateFlag="1" LoanNumber="APP-00099999">                                                             
        '    <LOANSTATUSES Operation="ADD" LoanStatusCode="PENDING COMPLETE" EffectiveDate="07/21/17"/>                
        '  </LOAN>                                                                                                     
        '</NLS>

        writeoffxml = NLSApp.SQLSelectStatementReadNext()

    Loop
    'Update last/Next Run datetime and Last user record
   ' Runstamp = NOW()
   ' NextRun = Date() + 1 & " 8:00:00 AM"
   ' NextRun = NOW()
    'UID = NLSApp.GetCurrentUserNo()

   ' Cifupdate = "<NLS><CIF UpdateFlag='1'CIFNumber='AUTO WRITEOFFCLOSE'><CIFDETAIL CIFDetail3='Next Run: " & NextRun & "' CIFDetail4='Last Run: " & Runstamp & "' CIFDetail5=' UserID: " & UID & "' /></CIF></NLS>"

   ' ImportXML(Cifupdate)

End Sub
'-------------------------------------------------------------------------------------------------------------------------

Sub AutoCloseaddCompleted()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Close Loans in COMPLETE PENDING status for prescribed waiting period or more and queue up ZERO Balance letter for printing or email
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    xmlsql = NLSapp.SQLSelectStatement("SELECT cd.userdef27 FROM cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'AUTO WRITEOFFCLOSE'")

    closurexml = NLSapp.SQLSelectStatement(xmlsql)

    If len(closurexml) = 0 Then
        closurexml = "***END OF RESULT***"
    End If

    Do Until closurexml = "***END OF RESULT***"
        refno = NLSapp.SQLSelectStatementNextColumn()
        template = NLSapp.SQLSelectStatementNextColumn()
        SendToQueue = NLSapp.SQLSelectStatementNextColumn()
        emailtoaddr = NLSapp.SQLSelectStatementNextColumn()
        name = NLSapp.SQLSelectStatementNextColumn()

       If PendingCompleteClosureImportXML(closurexml&refno) = "" Then

        'Expected XML string output Format:
        '<NLS>                                                                                                           
        '  <TRANSACTIONS>                                                                                                
        '    <TRANSACTIONCODE TransactionCode="462" EffectiveDate="07/21/17" LoanNumber="APP-00999999" Amount="537.43" />
        '  </TRANSACTIONS>                                                                                               
        '  <LOAN UpdateFlag="1" LoanNumber="APP-00999999" LoanStatusCode="CLOSED" ClosedDate="07/21/17" >                
        '    <LOANSTATUSES Operation="DELETE" LoanStatusCode="PENDING COMPLETE" EffectiveDate="07/21/17"/>               
        '    <LOANSTATUSES Operation="ADD" LoanStatusCode="COMPLETED" EffectiveDate="07/21/17"/>                         
        '  </LOAN>                                                                                                       
        '</NLS>


            Set DT = NLSApp.GetDocumentTemplates(template)
            DT.GenerateDocument refno, "TRUE", "SERVICING NOTES", SendToQueue, "T:\Temp\"

            If SendToQueue = "FALSE" Then
                If template = 5 Then
                    emailbody = "Dear " & name & chr(13) & _
                    "Your loan now has a $0 balance And the automatic ACH payments have been discontinued." & chr(13) & _
                    "We would Like to thank you for choosing Freedom Plus." & chr(13) & chr(13) & _
                    "Warmest regards," & chr(13) & _
                    "FreedomPlus Servicing" & chr(13) & _
                    "Phone: 800-297-5897"

                    NLSApp.SendEmailMessage "Servicing FPlus", emailtoaddr, "customersupport@freedomplus.com", "Your loan now has a $0 balance", emailbody 
                 End If
                If template = 6 Then
                    emailbody = "Dear " & name & chr(13) & _
                    "Your loan now has a $0 balance And the automatic ACH payments have been discontinued." & chr(13) & _
                    "We would Like to thank you for choosing Concolidation Plus." & chr(13) & chr(13) & _
                    "Warmest regards," & chr(13) & _
                    "Consolidation Plus Servicing" & chr(13) & _
                    "Phone: 1-888-950-4829"

                    NLSApp.SendEmailMessage "Servicing CPlus", emailtoaddr, "consolplus-servicing@consolplus.com", "Your loan now has a $0 balance", emailbody 
                 End If

            End If
        Else
            'PLACEHOLDER
            'NLSApp.MessageBox("There was a problem with refno " &  refno & ".")
        End If
        closurexml = NLSApp.SQLSelectStatementReadNext()
        If NLSApp.GetCurrentUserNo() = 39 Then
            'NLSApp.MessageBox(closurexml)
        End If
    Loop

    'RESET COMMENT userid to ADMINISTRATOR
    commentupdatesql = "Select 0; update loanacct_comments set created_by = 0,userdef01 = 'Outbound',userdef02 = 'Zero Balance Letter',userdef03 = 'Email' " & _
    "where comment_description in ('CPLUS ZERO BALANCE LETTER','FPLUS ZERO BALANCE LETTER') " & _
    "and created_by <> 0 and created >= left(getdate(),11)"
    
    nlsapp.sqlselectstatement(commentupdatesql)

    'Update last/Next Run datetime and Last user record
'    Runstamp = NOW()
'    NextRun = Date() + 1 & " 8:00:00 AM"
'    NextRun =NOW()
'    UID = NLSApp.GetCurrentUserNo()

'    Cifupdate = "<NLS><CIF UpdateFlag='1'CIFNumber='AUTO WRITEOFFCLOSE'><CIFDETAIL CIFDetail6='Next Run: " & NextRun & "' CIFDetail7='Last Run: " & Runstamp & "' CIFDetail8=' UserID: " & UID & "' /></CIF></NLS>"

'    ImportXML(Cifupdate)

End Sub

'-------------------------------------------------------------------------------------------------------------------------
Sub AutoReverseWO()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'For a Payoff NSF, Automaticly Reverse any Write-off transactions that happened after the pay-off
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    xmlsql = NLSapp.SQLSelectStatement("SELECT cd.userdef35 FROM cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'AUTO WRITEOFFCLOSE'")

    revwriteoffxml = NLSapp.SQLSelectStatement(xmlsql)

    If len(revwriteoffxml) = 0 Then
        revwriteoffxml = "***END OF RESULT***"
    End If

    Do Until revwriteoffxml = "***END OF RESULT***"

        ImportXML(revwriteoffxml)
        'Expected XML string output Format:
        '<NLS>                                                                                                         
        '  <TRANSACTIONS>                                                                                              
        '    <TRANSACTIONCODE TransactionCode="541" EffectiveDate="07/25/17" LoanNumber="APP-00614381" Amount="3.67" TransRefNo="4560197" />
        '    <TRANSACTIONCODE TransactionCode="545" EffectiveDate="07/25/17" LoanNumber="APP-00614381" Amount="3.67" TransRefNo="4560197" />
        '    <TRANSACTIONCODE TransactionCode="453" EffectiveDate="07/25/17" LoanNumber="APP-00614381" Amount="3.67" TransRefNo="4560197" />
        '    <TRANSACTIONCODE TransactionCode="449" EffectiveDate="07/25/17" LoanNumber="APP-00614381" Amount="3.67" TransRefNo="4560197" /> 
        '  </TRANSACTIONS>                       
        '</NLS>

        revwriteoffxml = NLSApp.SQLSelectStatementReadNext()

    Loop

End Sub
'-------------------------------------------------------------------------------------------------------------------------

Sub NonACHDetail2udf45Maint()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'For Updating Detail2 UDF45 with the ACH Company ID to be utilized by web portal one-tine ACH payment function
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    importsxml = NLSApp.SQLSelectStatement(NLSApp.SQLSelectStatement("SELECT cd.userdef35 FROM cif c, cif_detail cd Where c.cifno = cd.cifno and cifnumber = 'WEB PAY MAINT'"))
    If len(importsxml) = 0 Then
        importsxml = "***END OF RESULT***"
    End If

    Do Until importsxml = "***END OF RESULT***"

        ImportXML(importsxml)
        'Expected XML string output Format:
        '<NLS>                                                                                                           
        '  <LOAN  UpdateFlag = "1" AcctRefno="155190" >
        '   <LOANDETAIL2 UserDefined45 = "31" />
        '  </LOAN>
        '</NLS>

        importsxml = NLSApp.SQLSelectStatementReadNext()
    Loop
End Sub

'-------------------------------------------------------------------------------------------------------------------------
Sub replookup()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Adds list of debt settlemtn company representatives associated with selected company for LOAN_DETAIL2_UDF21 to drop down options
    'in LOAN_DETAIL2_UDF22.
    'Triggered when LOAN_DETAIL2_UDF21 field changes.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    currco = NLSApp.GetField("LOAN_DETAIL2_UDF21")
    repcountsql = "select count(row_id) from cif c, cif_cif_relationship ccr where c.cifno = ccr.cifno1 and c.company = '" & currco & "'"
    repcount = NLSapp.SQLSelectStatement(repcountsql)
    repnamesql = "select c2.shortname from cif c, cif_cif_relationship ccr, cif c2 " & _
              "where  c.cifno = ccr.cifno1 and c2.cifno = ccr.cifno2 and c.company = '" & currco & "'"
    indexcount = 0
    repname = NLSapp.SQLSelectStatement(repnamesql)     
    indexx =  NLSapp.SQLSelectStatementNextColumn()    

          
        If len(currco) > 0 Then
            If len(repname) = 0 Then
                NLSApp.MessageBox("The compamny, ' " & currco & "', currently has no representatives assinged to it." & chr(13)  & chr(13) & "Please select another company or go to CONTACTS to add a representative.")
                Exit Sub
            End If
        End If

    Do 
       NLSapp.AddComboString "LOAN_DETAIL2_UDF22", indexcount, repname

       repname = NLSapp.SQLSelectStatementReadNext()          
       indexx =  NLSapp.SQLSelectStatementNextColumn() 
       indexcount = indexcount + 1
    Loop While CDbl(indexcount) < CDbl(repcount)

End Sub

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Sub ACHCompanySyncatupdate()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Maintains sync with Class 2 (Investor) with ACH Company and detail 2 UDF45 (Web pay ACH company)
    'Triggered when individual loan account is chnaged or updated.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Declare variables
    runtimestamp = Now()
    emailyn = "N"
    userno = NLSApp.GetCurrentUserNo()
    refno = NLSApp.GetField("LOAN_REFNO")
    Investor = NLSApp.GetField("LOAN_CLASS2")
    Portfolio = NLSApp.GetField("LOAN_PORTFOLIO")
    WebACH = NLSApp.GetField("LOAN_DETAIL2_UDF45")
    targetACH =  InvestorACHSync(Investor, Portfolio)
    'SQL to look up all active ACH records that do not match target ACH Company ID
    nomatchACHsql = "select convert(varchar,row_id)  from loanacct_ach where status = 0 and acctrefno = " & refno & " and ach_company_id <> " & targetACH
    'Hold query results here
    nomatchACHrowID = NLSapp.SQLSelectStatement(nomatchACHsql) 
    
    'Other variables
    modhistupdatesql = ""  
    loanach = ""
    loandetail = ""
    achxml = ""
    Detail2pdatesql = ""
    ACHUpdatesql = ""
    
    'Did any ACH rows not match targetACH? 
    If Not nomatchACHrowID = "" Then
        Do Until nomatchACHrowID = "" '"***END OF RESULT***"
            ACHUpdatesql = "SELECT 0; UPDATE loanacct_ach SET ach_company_id = " & targetACH & " Where row_id = " & nomatchACHrowID
            NLSapp.SQLSelectStatement(ACHUpdatesql)
            nomatchACHrowID = NLSapp.SQLSelectStatement(nomatchACHsql)
        Loop

    Else
        'MsgBox("all match")
    End If    
    
    'Does the current Web Portal ACH Company ID match the correct ACH Company ID?
    If Not targetACH = WebACH Then
        Detail2updatesql = "SELECT 0; UPDATE loanacct_detail_2 SET userdef45 = '" & targetACH & "' Where acctrefno = " & refno
        NLSapp.SQLSelectStatement(Detail2updatesql)
    End If


    'If both ACH rows and WebACH match target ACH do nothing
'    If loanach = "" and loandetail = "" Then
        'MsgBox("no Update")
'    Else
        'Build xml import string, perform import and update history to show VBscript made the change
'        achxml = "<NLS EnforceTagExistence=""1"">" & chr(13) & _
'             "  <LOAN UpdateFlag=""1"" AcctRefno=""" & refno & """>" & chr(13) & loandetail & loanach & _
'             "  </LOAN>" & chr(13) & "</NLS>"  & refno & emailyn
        
        'XML Import           
'        ACHcorrectionNotice(achxml)
        
       'Update loan mod history to show ADMIN as user and add notation for VBScript for any loans processed
'        modhistupdatesql = "SELECT 0; UPDATE loanacct_mod_history SET mod_uid = 0, modification_description = '//VBScript//' +char(13) + modification_description where mod_uid = " & userno & " and (item_changed = 'ACH Setup'  or item_changed like '%Tab 2 Loan Userdef45%') and mod_datestamp >= '" & runtimestamp & "'"
'        NLSapp.SQLSelectStatement(modhistupdatesql)        
        
'    End If
      
End Sub

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Sub ACHCompanySyncatclose()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Maintains sync with Class 2 (Investor) with ACH Company and detail 2 UDF45 (Web pay ACH company)
    'Triggered by function to run script at system log out.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    emailyn = "N"
    userno = NLSApp.GetCurrentUserNo()
    runtimestamp = Now()
    'SQL query looks up out-of-sync loans
    refnosql = NLSapp.SQLSelectStatement("SELECT  userdef35 from  cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")

    'Has process made any emails the current day and for what loans?
    lastemailday = NLSapp.SQLSelectStatement("SELECT  userdef07 from  cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")

    'If  no emails the current day reset refno list
    If NOW() - CDate(lastemailday) > 1 Then
        NLSapp.SQLSelectStatement("SELECT 0; UPDATE cif_detail SET  userdef06 = ' ' From cif_detail cd, cif c where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")
    End If

    'If emails have been sent the current day, include the refno list in the sql statement
    lastrefnos = NLSapp.SQLSelectStatement("SELECT  userdef06 from  cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")
    If CDate(lastemailday) = DateValue(NOW()) Then
        acctfilter = "AND nonachUDFs.acctrefno NOT IN (" & lastrefnos & " 0)" & chr(13) & "order by acctrefno,[updateach] desc"
         RefNolist = lastrefnos
    Else
        acctfilter = "order by acctrefno,[updateach] desc"
    End If

    prevLoanRefNo = 0
    
    'Run SQL lookup
    LoanRefNo = NLSapp.SQLSelectStatement(refnosql & chr(13) & acctfilter)

    If Len(LoanRefNo) = 0 Then
        LoanRefNo = "***END OF RESULT***"
    End If

    Lcount = 0

    Do Until LoanRefNo = "***END OF RESULT***"
        RefNolist = RefNolist & LoanRefNo & ", "
        Lcount = Lcount + 1
        achid = NLSApp.SQLSelectStatementNextColumn()
        updateweb = NLSApp.SQLSelectStatementNextColumn()
        updateach = NLSApp.SQLSelectStatementNextColumn()
        achrowid = NLSApp.SQLSelectStatementNextColumn()

        'Never update loan detail twice
        If prevLoanRefNo = LoanRefNo Then
            loandetail = ""
'            emailyn = "Y"
        Else
            'If only loan detail is updated do not send email
            If updateweb = 1 Then
                loandetail = "    <LOANDETAIL2 UserDefined45=""" & achid & """/>" & chr(13)
'                emailyn = "N"
            End If

        End If
        'If ACH set up is updated send email
        If updateach = 1 Then
            loanach = "    <ACH Operation=""UPDATE"" ACHCompanyID=""" & achid & """ RowID=""" & achrowid & """/>" & chr(13)
'            emailyn = "Y"
        Else
            loanach = ""
        End If

        'Build xml import string
        If loandetail = "" And loanach = "" Then
            'NLSApp.MessageBox("No import")  
        Else

            achxml = "<NLS EnforceTagExistence=""1"">" & chr(13) & _
                    "  <LOAN UpdateFlag=""1"" AcctRefno=""" & LoanRefNo & """>" & chr(13) & loandetail & loanach & _
                    "  </LOAN>" & chr(13) & "</NLS>" & LoanRefNo & emailyn

            ACHcorrectionNotice(achxml)
        End If

        prevLoanRefNo = LoanRefNo

        'Update last email day and append current loan to refno list
        lastemailday = NLSapp.SQLSelectStatement("SELECT  userdef07 from  cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")
        'lastrefnos = NLSapp.SQLSelectStatement("SELECT  userdef06 from  cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")

        'Reset sgl statement for next loan
         acctfilter = "AND nonachUDFs.acctrefno NOT IN (" & RefNolist & " 0)" & chr(13) & "order by acctrefno,[updateach] desc"

        LoanRefNo = NLSapp.SQLSelectStatement(refnosql & chr(13) & acctfilter)

        If Len(LoanRefNo) = 0 Then
            LoanRefNo = "***END OF RESULT***"
        End If
    Loop

    'Update field that list refnos processed the current day
    newrefnolistsql = "SELECT 0; Update cif_detail SET userdef06 = '" & RefNolist & "', userdef07 = convert(char(8),getdate(),1) from cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'"
    NLSapp.SQLSelectStatement(newrefnolistsql)
 
    'Update loan mod history to show ADMIN as user and add notation for VBScript for any loans processed
    modhistupdatesql = "SELECT 0; UPDATE loanacct_mod_history SET mod_uid = 0, modification_description = '//VBScript//' +char(13) + modification_description where mod_uid = " & userno & " and (item_changed = 'ACH Setup'  or item_changed like '%Tab 2 Loan Userdef45%') and mod_datestamp >= '" & runtimestamp & "'"
    NLSapp.SQLSelectStatement(modhistupdatesql)
     
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GLOBAL FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Generic XML import funcxtion
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ImportXML(xml)
	NLSApp.ClearErrorString()

	If Not NLSApp.ImportXMLRecord(xml) Then
		If True Then
		NLSApp.MessageBox(xml & CHR(13) & NLSApp.GetErrorString())
		End If
	ImportXML = NLSApp.GetErrorString() & CHR(13) & CHR(13) & xml & CHR(13)
	Else
              ' NLSApp.MessageBox("Payment Added")
		ImportXML = ""
	End If
End Function
'-------------------------------------------------------------------------------------------------------------------------


Function ACHCoIDUDF45b(class2, portfolio)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ACH Company UDF Maintainance
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    achid = ""

    Select Case class2

        Case "1 - VULCAN"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "1"
            Else
                achid = "4"
            End If

        Case "3 - FFNAM FUND"
            achid = "3"
        Case "4 - RANGER"
            achid = "7"
        Case "5 - ATALAYA CAPITAL"
            achid = "8"
        Case "6 - RANGER DLT"
            achid = "9"
        Case "7 - BLUE ELEPHANT"
            achid = "10"
        Case "8 - TIDEPOOL"
            achid = "11"
        Case "9 - AEQUITAS ACC5"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "12"
            Else
                achid = "15"
            End If
        Case "5A - ATALAYA V-B"
            achid = "13"
        Case "5B - ATALAYA VI"
            achid = "14"
        Case "5C - ATALAYA MW"
            achid = "16"
        Case "7A - BLUE ELEPHANT OFFSHO"
            achid = "17"
        Case "5D - ATALAYA II-C"
            achid = "18"
        Case "10 - KEYSTONE"
            achid = "19"
        Case "11 - CRESTLINE"
            achid = "20"
        Case "12 - FCCF-SERIES 1"
            achid = "21"
        Case "13 - FCCF-SERIES 2"
            achid = "22"
        Case "14 - ACM III"
            achid = "23"
        Case "15 - ACM MW II"
            achid = "24"
        Case "16 - CRESTLINE II"
            achid = "25"
        Case "17 - DAKOY"
            achid = "26"
        Case "18 - STONERIDGE"
            achid = "27"
        Case "19 - ACM ALAMOSA I"
            achid = "32"
        Case "20 - ACM ALAMOSA I-A"
            achid = "28"
        Case "21 - ACM GATEWAY"
            achid = "30"
        Case "22 - ACM F ACQUISITION"
            achid = "31"
        Case "23 - 590 CONSUMER LENDING"
            achid = "29"
        Case "25 - ACM IV Off"
            achid = "33"
        Case "24 - ACM IV On"
            achid = "34"
        Case "13A - FCCF - CS"
            achid = "35"
        Case "13B-FCCF-ABS1"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "41"
            Else
                achid = "40"
            End If
        Case "13Z-FCCF-INEL"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "43"
            Else
                achid = "42"
            End If
    End Select

    ACHCoIDUDF45b = achid

End Function

Function PendingCompleteClosureImportXML(xml)
    NLSApp.ClearErrorString()
    endmark =  "</NLS>"
    refno = Mid(xml, InStr(xml, endmark) + len(endmark), len(xml) - InStr(xml, endmark) + len(endmark))
    
    If Not NLSApp.ImportXMLRecord(xml) Then
        If True Then
            errormsg = NLSApp.GetErrorString()
            subject = "Pending Complete Loan Closure process IMPORT FAILURE!"
            toaddress = "NLSErrors@freedomdebtrelief.com" '"somebody@bills.com" 
            body1 = "Acctrefno: " & refno & CHR(13) & "Import Error: " & chr(13) & errormsg & CHR(13) & CHR(13) & xml & CHR(13)
            NLSApp.SendEmailMessage "Loan Servicing3", toaddress, "svc.loanservicing@freedomdebtrelief.com", subject, body1
        End If
        PendingCompleteClosureImportXML = errormsg & CHR(13) & CHR(13) & xml & CHR(13)
    Else
        ' NLSApp.MessageBox("")
        PendingCompleteClosureImportXML = ""
    End If
End Function

Function Monthtext(monthval)
    mtext = ""
    Select Case monthval
        Case "1"
            mtext = "January"
        Case "2"
            mtext = "February"
        Case "3"
            mtext = "March"
        Case "4"
            mtext = "April"
        Case "5"
            mtext = "May"
        Case "6"
            mtext = "June"
        Case "7"
            mtext = "July"
        Case "8"
            mtext = "August"
        Case "9"
            mtext = "September"
        Case "10"
            mtext = "October"
        Case "11"
            mtext = "November"
        Case "12"
            mtext = "December"
        Case Else
            mtext = "Invalid Month value"
    End Select
    Monthtext = mtext
End Function

Function MonthNo(monthtxt)
    mthno = ""
    Select Case left(monthtxt, 3)
        Case "Jan"
            mthno = "1"
        Case "Feb"
            mthno = "2"
        Case "Mar"
            mthno = "3"
        Case "Apr"
            mthno = "4"
        Case "May"
            mthno = "5"
        Case "Jun"
            mthno = "6"
        Case "Jul"
            mthno = "7"
        Case "Aug"
            mthno = "8"
        Case "Sep"
            mthno = "9"
        Case "Oct"
            mthno = "10"
        Case "Nov"
            mthno = "11"
        Case "Dec"
            mthno = "12"
        Case Else
            mthno = "Invalid Month"
    End Select
    MonthNo = mthno
End Function

Function GetTotalFromtaskGrid101(strXMl, colname)
    strXMl = Replace(strXMl, "&", "&amp;")
    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    xmlDoc.setProperty "SelectionLanguage", "XPath"

    Err.Clear
    On Error Resume Next
    xmlDoc.loadxml(strXMl)
    If Err.Number <> 0 Then
        'Exit Function
        NLSapp.MessageBox(Err.Description)
    End If
    On Error GoTo 0
                
    Set nodesList = xmlDoc.documentelement.selectNodes("/DATA/ROW") 
    counter = nodesList.length

    dblTotal = 0

    For x = 1 To counter
                                        
        Set currNode=nodesList.nextNode 
        CurrNodeName = currNode.nodename
                                
        Set cnode = currNode.childnodes
                
        For Each c In cnode

            name = c.getAttribute("Name")
            value = c.getAttribute("Value")

            If name = colname Then
                dblTotal = dblTotal + value
            End If

        Next
    Next

    GetTotalFromtaskGrid101 = dblTotal / 2 
End Function

Function GetmonthsFromtaskGrid101(strXMl)
    strXMl = Replace(strXMl, "&", "&amp;")
    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    xmlDoc.setProperty "SelectionLanguage", "XPath"

    Err.Clear
    On Error Resume Next
    xmlDoc.loadxml(strXMl)
    If Err.Number <> 0 Then
        'Exit Function
        NLSapp.MessageBox(Err.Description)
    End If
    On Error GoTo 0
                
    Set nodesList = xmlDoc.documentelement.selectNodes("/DATA/ROW") 
    counter = nodesList.length

    GetmonthsFromtaskGrid101 = counter-1
End Function

Function POADocExists(refno)
    existynsql = "Select lc.row_id " & _
            "FROM loanacct_comments lc where lc.row_id in " & _
            "(" & _
            "Select  lcd.row_id " & _
            "from loanacct_comments lc, loanacct_comments_docs lcd " & _
            "where lc.row_id = lcd.row_id " & _
            "And lc.userdef02 = 'SETTLEMENT POA'               /* Power of Attorney */ " & _
            "and lc.category_id = 4                  /* Documents         */ " & _
            ") " & _
            "and lc.acctrefno = " & refno

    existyn = NLSApp.SQLselectstatement(existynsql)

    If len(existyn) > 0 Then
        POADocExists = "Y"
    Else
        POADocExists = "N"
    End If
End Function

Function InvestorACHSync(class2, portfolio)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Matches the class2 field (Investor)in a loan with the corresponding ACH company ID
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    achid = ""

    Select Case class2

        Case "1 - VULCAN"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "1"
            Else
                achid = "4"
            End If

        Case "3 - FFNAM FUND"
            achid = "3"
        Case "4 - RANGER"
            achid = "7"
        Case "5 - ATALAYA CAPITAL"
            achid = "8"
        Case "6 - RANGER DLT"
            achid = "9"
        Case "7 - BLUE ELEPHANT"
            achid = "10"
        Case "8 - TIDEPOOL"
            achid = "11"
        Case "9 - AEQUITAS ACC5"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "12"
            Else
                achid = "15"
            End If
        Case "5A - ATALAYA V-B"
            achid = "13"
        Case "5B - ATALAYA VI"
            achid = "14"
        Case "5C - ATALAYA MW"
            achid = "16"
        Case "7A - BLUE ELEPHANT OFFSHO"
            achid = "17"
        Case "5D - ATALAYA II-C"
            achid = "18"
        Case "10 - KEYSTONE"
            achid = "19"
        Case "11 - CRESTLINE"
            achid = "20"
        Case "12 - FCCF-SERIES 1"
            achid = "21"
        Case "13 - FCCF-SERIES 2"
            achid = "22"
        Case "14 - ACM III"
            achid = "23"
        Case "15 - ACM MW II"
            achid = "24"
        Case "16 - CRESTLINE II"
            achid = "25"
        Case "17 - DAKOY"
            achid = "26"
        Case "18 - STONERIDGE"
            achid = "27"
        Case "19 - ACM ALAMOSA I"
            achid = "32"
        Case "20 - ACM ALAMOSA I-A"
            achid = "28"
        Case "21 - ACM GATEWAY"
            achid = "30"
        Case "22 - ACM F ACQUISITION"
            achid = "31"
        Case "23 - 590 CONSUMER LENDING"
            achid = "29"
        Case "18A - STONE RIDGE II"
            achid = "34"
        Case "24 - ACM IV ON"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "36"
            Else
                achid = "37"
            End If            
        Case "25 - ACM IV OFF"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "35"
            Else
                achid = "38"
            End If
        Case "13A - FCCF - CS"
            achid = "39"
        Case "13B-FCCF-ABS1"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "41"
            Else
                achid = "40"
            End If
        Case "13Z-FCCF-INEL"
            If portfolio = "PLUS PORTFOLIO" Then
                achid = "43"
            Else
                achid = "42"
            End If
        Case "13C-FCCF-ABS2"
            achid = "45"
        Case Else
            If Not class2 = "" Then
                NLSApp.MessageBox("Cannot Locate Class2/ACH company ID for " & class2 & ".")
            End if
            achid = "9999"
    End Select

    InvestorACHSync = achid
    
'++++++++++++++++++++++++++++++++++++++++++++++++++++
'current ach Set up as of 11/29/17
'ach_company_id		ach_company_name
'0					NONE
'1					Vulcan F+
'2					Aequitas F+
'3					FFAM
'4					Vulcan C+
'5					Aequitas C+
'7					Ranger F+
'8					Atalaya
'9					Ranger DLT F+
'10					Blue Elephant F+
'11					TidePool F+
'12					Aequitas ACC5
'13					Atalaya V-B
'14					Atalaya VI
'15					AEQUITAS ACC5 C+
'16					Atalaya MW
'17					Blue Elephant Offshore F+
'18					Atalaya II-C
'19					KEYSTONE
'20					CRESTLINE
'21					FCCF-Series 1
'22					FCCF-Series 2
'23					ACM III
'24					ACM MW II
'25					CRESTLINE II
'26					DaKoy Income Fund LP
'27					StoneRidge
'28					ACM ALAMOSA I-A
'29					590 Consumer Lending Corp
'30					ACM Gateway
'31					ACM F Acquisition
'32					ACM Alamosa I
'33					Stone Ridge II C+
'34					Stone Ridge II F+
'35					ACM IV Off F+
'36					ACM IV On F+
'37					ACM IV On C+
'38					ACM IV Off C+
'39					13A - FCCF - CS

End Function

Function ACHcorrectionNotice(xml)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function processes the data updates to sync Class 2 (Investor) with ACH Company and detail 2 UDF45 (Web pay ACH company)
    ' as well as sending the email notification that an update has occured or if there is an error in processing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    NLSApp.ClearErrorString()
    'Get list of any accounts that have been processed by this script and sent an email the current day
    refnostodaysql = "SELECT distinct case when cd.userdef07 < convert(char(8),getdate(),1) then '0,' Else cd.userdef06 End  from cif c, cif_detail cd where c.cifno = cd.cifno  and c.cifnumber = 'ACH MONITOR' "
    refnostoday = left(NLSapp.SQLSelectStatement(refnostodaysql), 1)

    'Get Email body text ---
    bodyheada = NLSapp.SQLSelectStatement("select cd.userdef02 from cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")
    bodyhead1 = NLSapp.SQLSelectStatement("select cd.userdef03 from cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")
    bodyhead2 = NLSapp.SQLSelectStatement("select cd.userdef04 from cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")
    bodyhead3 = NLSapp.SQLSelectStatement("select cd.userdef05 from cif c, cif_detail cd where c.cifno = cd.cifno and c.cifnumber = 'ACH MONITOR'")
    endmark = "</NLS>"

    'Parse xml string passed to Func for  "Y" or "N" code to send an email for active loan and loanrefno ID
    emailyn = right(xml, 1)
    refnoyn = Mid(xml, InStr(xml, endmark) + len(endmark), len(xml) - InStr(xml, endmark) + len(endmark))
    refno = left(refnoyn, len(refnoyn) - 1)


    'Do not send an email if one has already been sent the current day
    If emailyn = "Y" Then
        'emailyn = NLSapp.SQLSelectStatement("SELECT case when '" & refno & "' NOT IN (" & Left(refnostoday, Len(refnostoday)-1) & ") then 'Y' else 'N' end")
        'msgbox("SELECT case when '" & refno & "' NOT IN (" & Left(refnostoday, Len(refnostoday)-1) & ") then 'Y' else 'N' end")
    End If

    'Perform Import
    If Not NLSApp.ImportXMLRecord(xml) Then

        If True Then
            'FOR IMPORT ERRORS ONLY and email should get sent
            errormsg = NLSApp.GetErrorString()
            emailyn = "Y"
        End If
        ACHcorrectionNotice = errormsg & CHR(13) & CHR(13) & xml & CHR(13)
        NLSApp.SendEmailMessage "Loan Servicing3", "NLSErrors@freedomdebtrelief.com", "svc.loanservicing@freedomdebtrelief.com", "XML Import ACH update ERROR", ACHcorrectionNotice
    Else

        ACHcorrectionNotice = ""
    End If

End Function
