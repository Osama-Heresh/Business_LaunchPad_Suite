import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } from 'docx';
import { saveAs } from 'file-saver';

export interface DocumentTemplate {
  id: string;
  title: string;
  description: string;
  category: string;
  downloadCount: number;
  fileSize: string;
  rating: number;
  tags: string[];
  preview?: string;
  createdAt: string;
  updatedAt: string;
  content?: string;
}

const templates: DocumentTemplate[] = [
  // Human Resources (HR) Templates
  {
    id: 'hr-offer-letter-fulltime',
    title: 'Employment Offer Letter (Full-Time)',
    description: 'Professional offer letter template for new employees with compensation details and terms',
    category: 'hr',
    downloadCount: 1250,
    fileSize: '43.9 KB',
    rating: 4.7,
    tags: ['employment', 'offer', 'full-time', 'compensation', 'benefits'],
    createdAt: '2024-01-15T10:00:00Z',
    updatedAt: '2024-01-20T14:30:00Z',
    content: `
[Company Letterhead]

[Date]

[Candidate Name]
[Candidate Address]
[City, State, ZIP Code]

Dear [Candidate Name],

We are pleased to offer you the position of [Job Title] at [Company Name]. We believe your skills and experience will be a valuable addition to our team.

POSITION DETAILS:
• Position: [Job Title]
• Department: [Department Name]
• Reports to: [Manager Name]
• Start Date: [Start Date]
• Employment Type: Full-Time

COMPENSATION & BENEFITS:
• Annual Salary: $[Salary Amount]
• Payment Schedule: [Bi-weekly/Monthly]
• Health Insurance: [Details]
• Retirement Plan: [401k/Other Details]
• Paid Time Off: [PTO Details]
• Other Benefits: [Additional Benefits]

TERMS & CONDITIONS:
• This offer is contingent upon successful completion of background check and reference verification
• Employment is at-will and may be terminated by either party with appropriate notice
• You will be required to sign our standard employment agreement and confidentiality agreement
• This position requires [any specific requirements]

Please confirm your acceptance of this offer by signing and returning this letter by [Response Deadline]. We look forward to welcoming you to our team.

Sincerely,

[Hiring Manager Name]
[Title]
[Company Name]

ACCEPTANCE:
I accept the terms of employment as outlined above.

Signature: _________________________ Date: _____________
[Candidate Name]
    `
  },
  {
    id: 'hr-employment-contract',
    title: 'Employment Contract Agreement',
    description: 'Comprehensive employment contract with detailed terms and legal provisions',
    category: 'hr',
    downloadCount: 980,
    fileSize: '67.2 KB',
    rating: 4.8,
    tags: ['contract', 'employment', 'legal', 'terms', 'agreement'],
    createdAt: '2024-01-10T09:00:00Z',
    updatedAt: '2024-01-25T16:45:00Z',
    content: `
EMPLOYMENT AGREEMENT

This Employment Agreement ("Agreement") is entered into on [Date] between [Company Name], a [State] corporation ("Company"), and [Employee Name] ("Employee").

1. POSITION AND DUTIES
Employee is hired as [Job Title] and will perform duties including:
• [Primary Responsibility 1]
• [Primary Responsibility 2]
• [Primary Responsibility 3]
• Other duties as assigned by management

2. TERM OF EMPLOYMENT
This agreement begins on [Start Date] and continues until terminated according to the terms herein.

3. COMPENSATION
• Base Salary: $[Amount] per year, paid [frequency]
• Performance Bonus: [Bonus Structure if applicable]
• Benefits: As outlined in Employee Handbook

4. WORKING HOURS
Standard work week is [Hours] hours, [Days] per week. Overtime compensation as per company policy and applicable law.

5. CONFIDENTIALITY
Employee agrees to maintain confidentiality of all proprietary information, trade secrets, and confidential business information.

6. NON-COMPETE (if applicable)
[Non-compete clause details]

7. TERMINATION
Either party may terminate this agreement with [Notice Period] written notice. Company may terminate immediately for cause.

8. GOVERNING LAW
This agreement is governed by the laws of [State].

IN WITNESS WHEREOF, the parties have executed this Agreement on the date first written above.

COMPANY:                           EMPLOYEE:

_________________________         _________________________
[Name], [Title]                   [Employee Name]
Date: _______________             Date: _______________
    `
  },
  {
    id: 'hr-nda-agreement',
    title: 'Non-Disclosure Agreement (NDA)',
    description: 'Professional NDA template to protect confidential information and trade secrets',
    category: 'hr',
    downloadCount: 2100,
    fileSize: '38.5 KB',
    rating: 4.6,
    tags: ['nda', 'confidentiality', 'legal', 'protection', 'trade-secrets'],
    createdAt: '2024-01-05T11:30:00Z',
    updatedAt: '2024-01-18T13:20:00Z',
    content: `
NON-DISCLOSURE AGREEMENT

This Non-Disclosure Agreement ("Agreement") is entered into on [Date] between [Company Name] ("Disclosing Party") and [Recipient Name] ("Receiving Party").

1. DEFINITION OF CONFIDENTIAL INFORMATION
Confidential Information includes all non-public information disclosed by the Disclosing Party, including but not limited to:
• Technical data, trade secrets, know-how
• Business plans, financial information
• Customer lists, supplier information
• Marketing strategies and plans
• Any other proprietary information

2. OBLIGATIONS OF RECEIVING PARTY
The Receiving Party agrees to:
• Hold all Confidential Information in strict confidence
• Not disclose Confidential Information to third parties
• Use Confidential Information solely for the purpose of [Purpose]
• Take reasonable precautions to protect confidentiality

3. EXCEPTIONS
This Agreement does not apply to information that:
• Is publicly available through no breach of this Agreement
• Was known to Receiving Party prior to disclosure
• Is independently developed without use of Confidential Information
• Is required to be disclosed by law

4. TERM
This Agreement remains in effect for [Duration] years from the date of signing, or until terminated by mutual consent.

5. RETURN OF MATERIALS
Upon termination, Receiving Party will return or destroy all materials containing Confidential Information.

6. REMEDIES
Breach of this Agreement may cause irreparable harm, and Disclosing Party may seek injunctive relief and monetary damages.

7. GOVERNING LAW
This Agreement is governed by the laws of [State/Country].

DISCLOSING PARTY:                  RECEIVING PARTY:

_________________________         _________________________
[Name], [Title]                   [Name]
[Company Name]                    
Date: _______________             Date: _______________
    `
  },
  {
    id: 'hr-employee-handbook',
    title: 'Employee Handbook',
    description: 'Comprehensive employee handbook covering policies, procedures, and company culture',
    category: 'hr',
    downloadCount: 756,
    fileSize: '125.8 KB',
    rating: 4.9,
    tags: ['handbook', 'policies', 'procedures', 'culture', 'guidelines'],
    createdAt: '2024-01-01T08:00:00Z',
    updatedAt: '2024-01-30T17:00:00Z',
    content: `
EMPLOYEE HANDBOOK
[Company Name]

TABLE OF CONTENTS
1. Welcome Message
2. Company Overview
3. Employment Policies
4. Compensation & Benefits
5. Time Off Policies
6. Workplace Conduct
7. Safety & Security
8. Communication Policies
9. Professional Development
10. Termination Procedures

1. WELCOME MESSAGE
Welcome to [Company Name]! This handbook contains important information about our company policies, procedures, and benefits.

2. COMPANY OVERVIEW
Mission: [Company Mission]
Vision: [Company Vision]
Values: [Company Values]

3. EMPLOYMENT POLICIES
• Equal Employment Opportunity
• At-Will Employment
• Background Checks
• Probationary Period: [Duration]

4. COMPENSATION & BENEFITS
• Payroll Schedule: [Schedule]
• Performance Reviews: [Frequency]
• Health Insurance: [Details]
• Retirement Plans: [Details]
• Other Benefits: [List]

5. TIME OFF POLICIES
• Paid Time Off (PTO): [Accrual Rate]
• Holidays: [List of Company Holidays]
• Sick Leave: [Policy Details]
• Family/Medical Leave: [FMLA Details]

6. WORKPLACE CONDUCT
• Professional Behavior Standards
• Dress Code: [Details]
• Anti-Harassment Policy
• Disciplinary Procedures

7. SAFETY & SECURITY
• Workplace Safety Guidelines
• Emergency Procedures
• Security Protocols
• Incident Reporting

8. COMMUNICATION POLICIES
• Email and Internet Usage
• Social Media Guidelines
• Confidentiality Requirements

9. PROFESSIONAL DEVELOPMENT
• Training Opportunities
• Tuition Reimbursement: [If applicable]
• Career Advancement

10. TERMINATION PROCEDURES
• Resignation Process
• Final Pay and Benefits
• Return of Company Property

This handbook is effective as of [Date] and supersedes all previous versions.

For questions, contact Human Resources at [Contact Information].
    `
  },
  {
    id: 'hr-job-description',
    title: 'Job Description Template',
    description: 'Structured template for creating comprehensive job descriptions with responsibilities and qualifications',
    category: 'hr',
    downloadCount: 1890,
    fileSize: '29.4 KB',
    rating: 4.5,
    tags: ['job-description', 'hiring', 'responsibilities', 'qualifications', 'requirements'],
    createdAt: '2024-01-12T12:00:00Z',
    updatedAt: '2024-01-22T10:15:00Z',
    content: `
JOB DESCRIPTION

POSITION INFORMATION
Job Title: [Job Title]
Department: [Department Name]
Reports To: [Manager Title]
Employment Type: [Full-time/Part-time/Contract]
Location: [Work Location]
Date: [Date Created/Updated]

JOB SUMMARY
[Brief 2-3 sentence overview of the role and its primary purpose]

KEY RESPONSIBILITIES
• [Primary responsibility 1 - be specific and action-oriented]
• [Primary responsibility 2]
• [Primary responsibility 3]
• [Primary responsibility 4]
• [Primary responsibility 5]
• [Additional responsibilities as assigned]

REQUIRED QUALIFICATIONS
Education:
• [Degree requirements]
• [Certifications if applicable]

Experience:
• [Years of experience required]
• [Specific industry experience]
• [Relevant work experience]

Skills & Competencies:
• [Technical skill 1]
• [Technical skill 2]
• [Soft skill 1]
• [Soft skill 2]
• [Software proficiency requirements]

PREFERRED QUALIFICATIONS
• [Additional education/certifications that would be beneficial]
• [Preferred experience]
• [Nice-to-have skills]

PHYSICAL REQUIREMENTS
• [Physical demands of the job]
• [Work environment conditions]
• [Travel requirements if any]

COMPENSATION & BENEFITS
• Salary Range: $[Min] - $[Max] annually
• [Benefits overview]
• [Performance incentives if applicable]

PERFORMANCE METRICS
Success in this role will be measured by:
• [Key performance indicator 1]
• [Key performance indicator 2]
• [Key performance indicator 3]

CAREER DEVELOPMENT
• [Growth opportunities]
• [Training provided]
• [Career progression path]

APPLICATION PROCESS
To apply, please submit:
• Resume
• Cover letter
• [Any additional requirements]

[Company Name] is an Equal Opportunity Employer committed to workplace diversity.

Approved by: _________________ Date: _________
[HR Manager Name]
    `
  },
  {
    id: 'hr-interview-evaluation',
    title: 'Interview Evaluation Form',
    description: 'Structured form to assess candidates consistently and fairly during interviews',
    category: 'hr',
    downloadCount: 1456,
    fileSize: '41.7 KB',
    rating: 4.4,
    tags: ['interview', 'evaluation', 'assessment', 'hiring', 'candidate'],
    createdAt: '2024-01-08T14:30:00Z',
    updatedAt: '2024-01-28T11:45:00Z',
    content: `
INTERVIEW EVALUATION FORM

CANDIDATE INFORMATION
Name: [Candidate Name]
Position Applied For: [Job Title]
Interview Date: [Date]
Interview Time: [Time]
Interviewer(s): [Interviewer Name(s)]
Interview Type: [Phone/Video/In-Person]

EVALUATION CRITERIA
Rate each area on a scale of 1-5 (1=Poor, 2=Below Average, 3=Average, 4=Good, 5=Excellent)

1. TECHNICAL SKILLS & QUALIFICATIONS
Education/Certifications: ___/5
Relevant Experience: ___/5
Technical Knowledge: ___/5
Industry Knowledge: ___/5
Comments: [Detailed feedback]

2. COMMUNICATION SKILLS
Verbal Communication: ___/5
Listening Skills: ___/5
Clarity of Expression: ___/5
Professional Presentation: ___/5
Comments: [Detailed feedback]

3. PROBLEM-SOLVING & ANALYTICAL THINKING
Problem-Solving Approach: ___/5
Critical Thinking: ___/5
Creativity/Innovation: ___/5
Decision-Making: ___/5
Comments: [Detailed feedback]

4. INTERPERSONAL SKILLS
Team Collaboration: ___/5
Leadership Potential: ___/5
Cultural Fit: ___/5
Adaptability: ___/5
Comments: [Detailed feedback]

5. MOTIVATION & ENTHUSIASM
Interest in Position: ___/5
Company Knowledge: ___/5
Career Goals Alignment: ___/5
Energy/Enthusiasm: ___/5
Comments: [Detailed feedback]

INTERVIEW QUESTIONS & RESPONSES
1. [Question 1]
Response: [Candidate's response and evaluation]

2. [Question 2]
Response: [Candidate's response and evaluation]

3. [Question 3]
Response: [Candidate's response and evaluation]

STRENGTHS
• [Key strength 1]
• [Key strength 2]
• [Key strength 3]

AREAS OF CONCERN
• [Concern 1]
• [Concern 2]
• [Concern 3]

OVERALL ASSESSMENT
Total Score: ___/100
Overall Rating: [Excellent/Good/Average/Below Average/Poor]

RECOMMENDATION
☐ Strongly Recommend for Hire
☐ Recommend for Hire
☐ Consider with Reservations
☐ Do Not Recommend
☐ Recommend for Different Position: [Specify]

NEXT STEPS
☐ Schedule Second Interview
☐ Check References
☐ Extend Offer
☐ Send Rejection
☐ Other: [Specify]

ADDITIONAL COMMENTS
[Any additional observations, concerns, or recommendations]

Interviewer Signature: _________________ Date: _________
[Interviewer Name]
    `
  },
  {
    id: 'hr-onboarding-checklist',
    title: 'Employee Onboarding Checklist',
    description: 'Comprehensive checklist to ensure smooth integration of new employees',
    category: 'hr',
    downloadCount: 2340,
    fileSize: '52.1 KB',
    rating: 4.8,
    tags: ['onboarding', 'checklist', 'integration', 'new-employee', 'orientation'],
    createdAt: '2024-01-03T09:15:00Z',
    updatedAt: '2024-01-26T15:30:00Z',
    content: `
EMPLOYEE ONBOARDING CHECKLIST

EMPLOYEE INFORMATION
Name: [Employee Name]
Position: [Job Title]
Department: [Department]
Start Date: [Start Date]
Manager: [Manager Name]
HR Representative: [HR Rep Name]

PRE-ARRIVAL PREPARATION (HR)
☐ Prepare employment paperwork
☐ Set up workstation/office space
☐ Order business cards
☐ Create email account
☐ Set up computer/equipment
☐ Prepare welcome packet
☐ Schedule first-day meetings
☐ Notify team of new hire
☐ Prepare ID badge/access cards
☐ Complete: _______ Date: _______

FIRST DAY ACTIVITIES
☐ Welcome and office tour
☐ Introduction to team members
☐ Review job description and expectations
☐ Complete I-9 verification
☐ Complete tax forms (W-4)
☐ Enroll in benefits programs
☐ Review employee handbook
☐ Set up direct deposit
☐ Provide company directory
☐ Issue equipment and supplies
☐ Complete: _______ Date: _______

FIRST WEEK ACTIVITIES
☐ Department orientation
☐ Meet with direct supervisor
☐ Review company policies
☐ Set up necessary accounts/systems
☐ Begin job-specific training
☐ Assign initial projects/tasks
☐ Schedule regular check-ins
☐ Introduce to key stakeholders
☐ Review performance expectations
☐ Complete safety training
☐ Complete: _______ Date: _______

FIRST MONTH ACTIVITIES
☐ Complete job-specific training
☐ 30-day performance check-in
☐ Address any questions/concerns
☐ Review goals and objectives
☐ Gather feedback on onboarding process
☐ Ensure system access is working
☐ Complete compliance training
☐ Schedule ongoing training sessions
☐ Review probationary period expectations
☐ Complete: _______ Date: _______

REQUIRED DOCUMENTATION
☐ Signed offer letter
☐ I-9 form with supporting documents
☐ W-4 tax form
☐ Emergency contact information
☐ Direct deposit authorization
☐ Benefits enrollment forms
☐ Employee handbook acknowledgment
☐ Confidentiality agreement
☐ IT acceptable use policy
☐ Safety training completion
☐ Complete: _______ Date: _______

EQUIPMENT & ACCESS
☐ Computer/laptop
☐ Phone/mobile device
☐ ID badge/access card
☐ Parking pass
☐ Office keys
☐ Email account setup
☐ System access permissions
☐ Software licenses
☐ Office supplies
☐ Business cards
☐ Complete: _______ Date: _______

TRAINING SCHEDULE
Week 1: [Training topics]
Week 2: [Training topics]
Week 3: [Training topics]
Week 4: [Training topics]

FEEDBACK & EVALUATION
30-Day Review Date: [Date]
60-Day Review Date: [Date]
90-Day Review Date: [Date]

Employee Feedback:
[Space for employee comments on onboarding experience]

HR Representative: _________________ Date: _______
Manager: _________________ Date: _______
Employee: _________________ Date: _______
    `
  },
  {
    id: 'hr-timesheet-template',
    title: 'Timesheet Template',
    description: 'Professional timesheet for tracking work hours and project time',
    category: 'hr',
    downloadCount: 3200,
    fileSize: '35.6 KB',
    rating: 4.3,
    tags: ['timesheet', 'hours', 'tracking', 'payroll', 'time-management'],
    createdAt: '2024-01-07T13:00:00Z',
    updatedAt: '2024-01-21T09:30:00Z',
    content: `
EMPLOYEE TIMESHEET

EMPLOYEE INFORMATION
Name: [Employee Name]
Employee ID: [ID Number]
Department: [Department]
Manager: [Manager Name]
Pay Period: [Start Date] to [End Date]

DAILY TIME RECORD
[Create a table format for daily entries]

DATE | DAY | START TIME | LUNCH OUT | LUNCH IN | END TIME | REGULAR HOURS | OVERTIME HOURS | TOTAL HOURS | PROJECT/TASK CODE
[Date] | Mon | [Time] | [Time] | [Time] | [Time] | [Hours] | [Hours] | [Hours] | [Code]
[Date] | Tue | [Time] | [Time] | [Time] | [Time] | [Hours] | [Hours] | [Hours] | [Code]
[Date] | Wed | [Time] | [Time] | [Time] | [Time] | [Hours] | [Hours] | [Hours] | [Code]
[Date] | Thu | [Time] | [Time] | [Time] | [Time] | [Hours] | [Hours] | [Hours] | [Code]
[Date] | Fri | [Time] | [Time] | [Time] | [Time] | [Hours] | [Hours] | [Hours] | [Code]
[Date] | Sat | [Time] | [Time] | [Time] | [Time] | [Hours] | [Hours] | [Hours] | [Code]
[Date] | Sun | [Time] | [Time] | [Time] | [Time] | [Hours] | [Hours] | [Hours] | [Code]

WEEKLY TOTALS
Regular Hours: [Total Regular Hours]
Overtime Hours: [Total Overtime Hours]
Total Hours: [Total Hours Worked]

PROJECT/TASK BREAKDOWN
Project Code: [Code] | Description: [Description] | Hours: [Hours]
Project Code: [Code] | Description: [Description] | Hours: [Hours]
Project Code: [Code] | Description: [Description] | Hours: [Hours]

TIME OFF USED
Vacation Hours: [Hours]
Sick Hours: [Hours]
Personal Hours: [Hours]
Holiday Hours: [Hours]

EXPENSE REIMBURSEMENTS (if applicable)
Date: [Date] | Description: [Description] | Amount: $[Amount]
Date: [Date] | Description: [Description] | Amount: $[Amount]
Total Expenses: $[Total Amount]

EMPLOYEE CERTIFICATION
I certify that the information contained in this timesheet is accurate and complete to the best of my knowledge.

Employee Signature: _________________ Date: _______

SUPERVISOR APPROVAL
I have reviewed and approve this timesheet for payroll processing.

Supervisor Signature: _________________ Date: _______
Supervisor Name: [Print Name]

HR USE ONLY
Processed by: [HR Rep Name]
Date Processed: [Date]
Payroll Period: [Period]
Notes: [Any special notes]
    `
  },
  {
    id: 'hr-leave-request-form',
    title: 'Leave/Time-Off Request Form',
    description: 'Formal request form for various types of employee leave and time off',
    category: 'hr',
    downloadCount: 2890,
    fileSize: '33.2 KB',
    rating: 4.6,
    tags: ['leave', 'time-off', 'request', 'vacation', 'pto'],
    createdAt: '2024-01-09T10:45:00Z',
    updatedAt: '2024-01-24T14:20:00Z',
    content: `
LEAVE/TIME-OFF REQUEST FORM

EMPLOYEE INFORMATION
Name: [Employee Name]
Employee ID: [ID Number]
Department: [Department]
Position: [Job Title]
Manager: [Manager Name]
Date of Request: [Date]

LEAVE REQUEST DETAILS
Type of Leave Requested:
☐ Vacation/PTO
☐ Sick Leave
☐ Personal Leave
☐ Family/Medical Leave (FMLA)
☐ Bereavement Leave
☐ Jury Duty
☐ Military Leave
☐ Other: [Specify]

DATES REQUESTED
Start Date: [Date]
End Date: [Date]
Total Days/Hours Requested: [Amount]
Return to Work Date: [Date]

Will this be:
☐ Consecutive days off
☐ Intermittent leave
☐ Reduced schedule

If intermittent or reduced schedule, please specify:
[Details of schedule]

REASON FOR LEAVE
[Provide brief explanation - medical details not required]

CURRENT LEAVE BALANCES (HR Use)
Vacation/PTO Balance: [Hours]
Sick Leave Balance: [Hours]
Personal Leave Balance: [Hours]
Other: [Hours]

WORK COVERAGE ARRANGEMENTS
Who will cover your responsibilities?
Name: [Coverage Person]
Contact: [Phone/Email]

Have you briefed your coverage person? ☐ Yes ☐ No

Key projects/deadlines during absence:
[List important items]

ADDITIONAL INFORMATION
Is this leave related to a workers' compensation claim? ☐ Yes ☐ No
If yes, claim number: [Number]

Have you taken leave for this same condition in the past 12 months? ☐ Yes ☐ No
If yes, dates: [Dates]

REQUIRED DOCUMENTATION
For medical leave, attach:
☐ Medical certification
☐ Fitness for duty certification (for return)

For family leave, attach:
☐ Birth certificate
☐ Medical certification
☐ Other: [Specify]

EMPLOYEE ACKNOWLEDGMENT
I understand that:
• This request is subject to approval
• Medical certification may be required
• I must follow company call-in procedures
• Failure to return on scheduled date may result in disciplinary action
• Leave may be unpaid if insufficient accrued time

Employee Signature: _________________ Date: _______

SUPERVISOR APPROVAL
☐ Approved
☐ Approved with modifications: [Specify]
☐ Denied - Reason: [Reason]

Supervisor Signature: _________________ Date: _______
Supervisor Name: [Print Name]

HR APPROVAL
☐ Approved
☐ Requires additional documentation
☐ FMLA eligible - paperwork sent
☐ Other: [Notes]

HR Representative: _________________ Date: _______
HR Name: [Print Name]

RETURN TO WORK
Actual Return Date: [Date]
Fitness for Duty Certification Received: ☐ Yes ☐ No ☐ N/A

HR Representative: _________________ Date: _______
    `
  },
  {
    id: 'hr-warning-letter',
    title: 'Employee Warning Letter',
    description: 'Formal disciplinary warning letter for performance or conduct issues',
    category: 'hr',
    downloadCount: 1670,
    fileSize: '28.9 KB',
    rating: 4.2,
    tags: ['warning', 'disciplinary', 'performance', 'conduct', 'formal'],
    createdAt: '2024-01-11T11:20:00Z',
    updatedAt: '2024-01-27T16:10:00Z',
    content: `
EMPLOYEE WARNING LETTER

[Company Letterhead]

Date: [Date]

TO: [Employee Name]
FROM: [Manager Name], [Title]
CC: Human Resources
RE: [Disciplinary Action Type] Warning

This letter serves as a [VERBAL/WRITTEN/FINAL] warning regarding your [performance/conduct] as an employee of [Company Name].

ISSUE DESCRIPTION
The following issue(s) have been identified:
[Detailed description of the problem, including specific incidents, dates, and behaviors]

COMPANY POLICY/STANDARD
This behavior/performance violates the following company policy or standard:
[Reference specific policy from employee handbook or job description]

PREVIOUS DISCUSSIONS
☐ This is the first documented discussion of this issue
☐ Previous discussions occurred on: [Date(s)]
☐ Previous warnings issued: [Details]

EXPECTED IMPROVEMENT
You are expected to immediately:
• [Specific action item 1]
• [Specific action item 2]
• [Specific action item 3]
• [Additional requirements]

CONSEQUENCES
Failure to improve your [performance/conduct] may result in:
☐ Additional disciplinary action
☐ Suspension without pay
☐ Termination of employment
☐ Other: [Specify]

IMPROVEMENT TIMELINE
You have [time period] to demonstrate sustained improvement. Your progress will be reviewed on [review date].

SUPPORT AVAILABLE
The company will provide the following support to help you succeed:
• [Training opportunities]
• [Resources available]
• [Mentoring/coaching]
• [Other support]

EMPLOYEE RESPONSE
You have the right to respond to this warning. Please provide your written response within [time period] if you wish to do so.

Employee Response:
[Space for employee to write response]

ACKNOWLEDGMENT
By signing below, you acknowledge that:
• You have received and read this warning
• You understand the issues described
• You understand the consequences of not improving
• You understand the timeline for improvement
• Signing does not necessarily indicate agreement

This warning will be placed in your personnel file.

Employee Signature: _________________ Date: _______
Employee Name (Print): [Employee Name]

Manager Signature: _________________ Date: _______
Manager Name (Print): [Manager Name]

HR Representative: _________________ Date: _______
HR Name (Print): [HR Rep Name]

Witness (if applicable): _________________ Date: _______
Witness Name (Print): [Witness Name]
    `
  },
  {
    id: 'hr-resignation-acceptance',
    title: 'Resignation Acceptance Letter',
    description: 'Professional letter for accepting employee resignations and outlining transition',
    category: 'hr',
    downloadCount: 1230,
    fileSize: '31.4 KB',
    rating: 4.5,
    tags: ['resignation', 'acceptance', 'transition', 'departure', 'formal'],
    createdAt: '2024-01-13T15:45:00Z',
    updatedAt: '2024-01-29T12:00:00Z',
    content: `
RESIGNATION ACCEPTANCE LETTER

[Company Letterhead]

Date: [Date]

[Employee Name]
[Employee Address]
[City, State, ZIP Code]

Dear [Employee Name],

This letter acknowledges receipt of your resignation letter dated [Date of Resignation Letter], in which you indicated your intention to resign from your position as [Job Title] with [Company Name].

RESIGNATION ACCEPTANCE
We accept your resignation effective [Last Working Date]. Your final day of employment will be [Last Working Date].

TRANSITION PERIOD
During your remaining time with the company, we ask that you:
• Complete current projects or transition them to [Successor/Team]
• Document ongoing processes and procedures
• Train [Replacement Name] or team members as needed
• Return all company property (listed below)
• Maintain confidentiality of all proprietary information

FINAL COMPENSATION
Your final paycheck will include:
• Salary through your last day of work
• Payment for unused vacation time: [Amount/Days]
• [Other applicable compensation]

Final pay will be processed according to state law and company policy and will be available on [Date].

BENEFITS CONTINUATION
• Health insurance coverage will end on [Date]
• COBRA continuation coverage information will be mailed separately
• 401(k) plan information and options will be provided by [Provider]
• [Other benefit details]

COMPANY PROPERTY RETURN
Please return the following items by your last day:
☐ Company laptop/computer
☐ Mobile phone/devices
☐ ID badge/access cards
☐ Keys
☐ Company credit cards
☐ Files and documents
☐ [Other company property]

EXIT INTERVIEW
An exit interview has been scheduled for [Date] at [Time] with [HR Representative]. This is an opportunity to provide feedback about your experience with the company.

REFERENCES
[Company Name] will provide employment verification and, upon request, a reference letter confirming your employment dates, position, and general performance.

NON-DISCLOSURE REMINDER
Please remember that your confidentiality obligations continue after your employment ends. You are required to maintain the confidentiality of all proprietary information, trade secrets, and confidential business information.

FUTURE OPPORTUNITIES
We wish you success in your future endeavors. Should circumstances change, we would welcome the opportunity to discuss potential future employment.

Thank you for your contributions to [Company Name]. We appreciate your [time period] of service and wish you all the best in your new position.

Sincerely,

[Manager Name]
[Title]
[Company Name]

CC: Human Resources
    Personnel File

EMPLOYEE ACKNOWLEDGMENT
I acknowledge receipt of this resignation acceptance letter and understand the terms outlined above.

Employee Signature: _________________ Date: _______
[Employee Name]
    `
  },
  {
    id: 'hr-exit-interview',
    title: 'Exit Interview Form',
    description: 'Comprehensive form to gather valuable feedback from departing employees',
    category: 'hr',
    downloadCount: 1890,
    fileSize: '47.3 KB',
    rating: 4.7,
    tags: ['exit-interview', 'feedback', 'departure', 'evaluation', 'improvement'],
    createdAt: '2024-01-14T08:30:00Z',
    updatedAt: '2024-01-31T10:45:00Z',
    content: `
EXIT INTERVIEW FORM

EMPLOYEE INFORMATION
Name: [Employee Name]
Position: [Job Title]
Department: [Department]
Manager: [Manager Name]
Hire Date: [Date]
Last Working Date: [Date]
Length of Service: [Years/Months]

INTERVIEW DETAILS
Interview Date: [Date]
Interviewer: [HR Rep Name]
Interview Type: ☐ In-Person ☐ Phone ☐ Video

REASON FOR LEAVING
Primary reason for leaving:
☐ Better opportunity/career advancement
☐ Higher compensation
☐ Relocation
☐ Work-life balance
☐ Management/supervision issues
☐ Company culture
☐ Lack of growth opportunities
☐ Job dissatisfaction
☐ Personal reasons
☐ Retirement
☐ Other: [Specify]

Please elaborate on your reason for leaving:
[Detailed response]

JOB SATISFACTION
Rate your overall job satisfaction (1=Very Dissatisfied, 5=Very Satisfied):

Job responsibilities: 1  2  3  4  5
Workload: 1  2  3  4  5
Work-life balance: 1  2  3  4  5
Compensation: 1  2  3  4  5
Benefits: 1  2  3  4  5
Recognition: 1  2  3  4  5
Training/development: 1  2  3  4  5
Career advancement: 1  2  3  4  5

MANAGEMENT & SUPERVISION
Rate your direct supervisor (1=Poor, 5=Excellent):

Communication: 1  2  3  4  5
Support: 1  2  3  4  5
Feedback: 1  2  3  4  5
Fairness: 1  2  3  4  5
Leadership: 1  2  3  4  5

Comments about management:
[Response]

COMPANY CULTURE
Rate the following aspects (1=Poor, 5=Excellent):

Teamwork: 1  2  3  4  5
Communication: 1  2  3  4  5
Company values: 1  2  3  4  5
Diversity & inclusion: 1  2  3  4  5
Work environment: 1  2  3  4  5

What did you like most about working here?
[Response]

What did you like least about working here?
[Response]

TRAINING & DEVELOPMENT
Did you receive adequate training for your role? ☐ Yes ☐ No
Were professional development opportunities available? ☐ Yes ☐ No
Did you feel supported in your career growth? ☐ Yes ☐ No

Comments:
[Response]

RECOMMENDATIONS FOR IMPROVEMENT
What could the company do to improve:

Employee retention:
[Response]

Management practices:
[Response]

Work environment:
[Response]

Training/development:
[Response]

Communication:
[Response]

WOULD YOU RECOMMEND?
Would you recommend [Company Name] as a good place to work? ☐ Yes ☐ No
Why or why not?
[Response]

Would you consider returning to work here in the future? ☐ Yes ☐ No ☐ Maybe

ADDITIONAL COMMENTS
Is there anything else you'd like to share about your experience?
[Response]

CONFIDENTIALITY REMINDER
☐ Employee reminded of ongoing confidentiality obligations
☐ Non-compete agreement discussed (if applicable)
☐ Return of company property confirmed

INTERVIEWER NOTES
[Space for HR representative's observations and notes]

Employee Signature: _________________ Date: _______

HR Representative: _________________ Date: _______

Thank you for your time and feedback. This information will be used to improve our workplace for current and future employees.
    `
  },
  {
    id: 'hr-termination-letter',
    title: 'Termination Letter',
    description: 'Professional termination letter with proper documentation and legal compliance',
    category: 'hr',
    downloadCount: 890,
    fileSize: '39.7 KB',
    rating: 4.1,
    tags: ['termination', 'employment', 'legal', 'documentation', 'formal'],
    createdAt: '2024-01-16T09:00:00Z',
    updatedAt: '2024-01-30T13:15:00Z',
    content: `
EMPLOYMENT TERMINATION LETTER

[Company Letterhead]

Date: [Date]

[Employee Name]
[Employee Address]
[City, State, ZIP Code]

Dear [Employee Name],

This letter serves as formal notification that your employment with [Company Name] is being terminated effective [Termination Date].

REASON FOR TERMINATION
☐ Voluntary resignation
☐ Position elimination/layoff
☐ Performance issues
☐ Violation of company policy
☐ Misconduct
☐ End of contract/temporary assignment
☐ Other: [Specify reason]

[Detailed explanation of termination reason - be factual and objective]

FINAL WORK DATE
Your last day of work is [Date]. You are [required/not required] to work through this date.

☐ You are relieved of duties immediately
☐ You will work through your final date
☐ You are on paid administrative leave until termination date

FINAL COMPENSATION
Your final paycheck will include:
• Salary/wages through [Date]: $[Amount]
• Unused vacation time: [Days] = $[Amount]
• [Other compensation items]
• Total final pay: $[Amount]

Final payment will be available on [Date] via [direct deposit/check].

BENEFITS INFORMATION
• Health insurance ends: [Date]
• COBRA information will be mailed within 14 days
• 401(k) plan: Contact [Provider] at [Phone] for options
• Life insurance: Coverage ends [Date]
• [Other benefit details]

COMPANY PROPERTY
You must return all company property by [Date]:
☐ Laptop/computer equipment
☐ Mobile phone/devices
☐ ID badge/access cards
☐ Keys
☐ Company vehicle
☐ Credit cards
☐ Files/documents
☐ [Other items]

Failure to return company property may result in deduction from final pay as permitted by law.

CONFIDENTIALITY & NON-COMPETE
You are reminded that your obligations regarding:
• Confidentiality of proprietary information
• Non-disclosure of trade secrets
• [Non-compete agreement terms, if applicable]
• Non-solicitation of employees/customers

These obligations continue after termination as outlined in your employment agreement.

REFERENCES
[Company Name] will verify employment dates, position, and salary information. [Additional reference policy information]

UNEMPLOYMENT BENEFITS
You may be eligible for unemployment benefits. Contact your state unemployment office for information and to file a claim.

APPEAL PROCESS (if applicable)
If you believe this termination is unjust, you may appeal by [process details] within [time period].

We wish you success in your future endeavors.

Sincerely,

[Manager/HR Name]
[Title]
[Company Name]

RECEIPT ACKNOWLEDGMENT
I acknowledge receipt of this termination letter and understand the terms outlined above.

Employee Signature: _________________ Date: _______
[Employee Name]

Witness: _________________ Date: _______
[Witness Name, Title]
    `
  },
  {
    id: 'hr-payroll-authorization',
    title: 'Payroll Authorization Form',
    description: 'Authorization form for payroll deductions and direct deposit setup',
    category: 'hr',
    downloadCount: 2150,
    fileSize: '42.8 KB',
    rating: 4.4,
    tags: ['payroll', 'authorization', 'deductions', 'direct-deposit', 'benefits'],
    createdAt: '2024-01-17T12:30:00Z',
    updatedAt: '2024-01-28T08:45:00Z',
    content: `
PAYROLL AUTHORIZATION FORM

EMPLOYEE INFORMATION
Name: [Employee Name]
Employee ID: [ID Number]
Social Security Number: XXX-XX-[Last 4 digits]
Department: [Department]
Position: [Job Title]
Hire Date: [Date]

DIRECT DEPOSIT AUTHORIZATION
I authorize [Company Name] to deposit my pay directly into my account(s) as indicated below:

PRIMARY ACCOUNT
☐ Checking ☐ Savings
Bank Name: [Bank Name]
Routing Number: [9-digit routing number]
Account Number: [Account number]
Amount: ☐ Entire paycheck ☐ $[Amount] ☐ [Percentage]%

SECONDARY ACCOUNT (Optional)
☐ Checking ☐ Savings
Bank Name: [Bank Name]
Routing Number: [9-digit routing number]
Account Number: [Account number]
Amount: ☐ Remaining balance ☐ $[Amount] ☐ [Percentage]%

PAYROLL DEDUCTIONS AUTHORIZATION
I authorize the following deductions from my pay:

REQUIRED DEDUCTIONS
☐ Federal income tax
☐ State income tax
☐ Social Security tax
☐ Medicare tax
☐ State disability insurance (if applicable)
☐ Unemployment insurance (if applicable)

VOLUNTARY DEDUCTIONS
☐ Health insurance premium: $[Amount] per pay period
☐ Dental insurance premium: $[Amount] per pay period
☐ Vision insurance premium: $[Amount] per pay period
☐ Life insurance premium: $[Amount] per pay period
☐ 401(k) contribution: [Percentage]% or $[Amount] per pay period
☐ Flexible Spending Account: $[Amount] per pay period
☐ Health Savings Account: $[Amount] per pay period
☐ Parking fees: $[Amount] per pay period
☐ Union dues: $[Amount] per pay period
☐ Charitable contributions: $[Amount] per pay period to [Organization]
☐ Other: [Description] $[Amount] per pay period

GARNISHMENTS/COURT ORDERS
☐ Child support: $[Amount] per pay period
☐ Wage garnishment: $[Amount] per pay period
☐ Tax levy: $[Amount] per pay period
☐ Other court-ordered deduction: [Description] $[Amount]

PAY SCHEDULE INFORMATION
Pay frequency: ☐ Weekly ☐ Bi-weekly ☐ Semi-monthly ☐ Monthly
Pay dates: [Specify pay dates]
First paycheck date: [Date]

EMERGENCY CONTACT FOR PAYROLL
Name: [Contact Name]
Relationship: [Relationship]
Phone: [Phone Number]
Email: [Email Address]

AUTHORIZATION & ACKNOWLEDGMENT
I understand and agree to the following:
• Deductions will continue each pay period until I submit a written change request
• Changes to deductions require [notice period] advance notice
• Direct deposit may take 1-2 pay periods to become effective
• I will receive pay stubs showing all deductions
• I am responsible for notifying payroll of any bank account changes
• The company is not responsible for delays caused by incorrect banking information

EMPLOYEE AUTHORIZATION
I authorize [Company Name] to make the deductions specified above and to deposit my net pay as indicated. I understand this authorization will remain in effect until I provide written notice of changes.

Employee Signature: _________________ Date: _______
Employee Name (Print): [Employee Name]

PAYROLL DEPARTMENT USE ONLY
Processed by: [Payroll Rep Name]
Date processed: [Date]
Effective pay period: [Date]
System updated: ☐ Yes
Notes: [Any special instructions]

Payroll Representative: _________________ Date: _______

HR APPROVAL
HR Representative: _________________ Date: _______
HR Name (Print): [HR Rep Name]
    `
  }
];

// Finance & Accounting Templates
const financeTemplates: DocumentTemplate[] = [
  {
    id: 'finance-invoice-template',
    title: 'Invoice Template',
    description: 'Professional invoice template for billing clients with itemized services and payment terms',
    category: 'finance',
    downloadCount: 3450,
    fileSize: '45.2 KB',
    rating: 4.8,
    tags: ['invoice', 'billing', 'payment', 'client', 'accounting'],
    createdAt: '2024-01-05T09:00:00Z',
    updatedAt: '2024-01-25T14:30:00Z',
    content: `
INVOICE

[Company Name]
[Company Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]
[Website]

BILL TO:
[Client Name]
[Client Company]
[Client Address]
[City, State, ZIP Code]

INVOICE DETAILS:
Invoice Number: [Invoice Number]
Invoice Date: [Date]
Due Date: [Due Date]
Payment Terms: [Net 30/Net 15/Due on Receipt]

DESCRIPTION OF SERVICES/PRODUCTS:

Item | Description | Quantity | Rate | Amount
-----|-------------|----------|------|-------
[Item 1] | [Description] | [Qty] | $[Rate] | $[Amount]
[Item 2] | [Description] | [Qty] | $[Rate] | $[Amount]
[Item 3] | [Description] | [Qty] | $[Rate] | $[Amount]

SUMMARY:
Subtotal: $[Subtotal]
Tax ([Tax Rate]%): $[Tax Amount]
Discount: $[Discount Amount]
TOTAL AMOUNT DUE: $[Total Amount]

PAYMENT INFORMATION:
Payment Method: [Check/Bank Transfer/Credit Card/PayPal]

Bank Details (if applicable):
Bank Name: [Bank Name]
Account Number: [Account Number]
Routing Number: [Routing Number]

TERMS & CONDITIONS:
• Payment is due within [Payment Terms] of invoice date
• Late payments may incur a [Late Fee]% monthly service charge
• Please include invoice number with payment
• All work performed and materials supplied are subject to our standard terms and conditions

NOTES:
[Additional notes or special instructions]

Thank you for your business!

[Authorized Signature]
[Name], [Title]
[Date]
    `
  },
  {
    id: 'finance-receipt-template',
    title: 'Receipt Template',
    description: 'Professional receipt template for documenting payments and transactions',
    category: 'finance',
    downloadCount: 2890,
    fileSize: '32.1 KB',
    rating: 4.6,
    tags: ['receipt', 'payment', 'transaction', 'proof', 'accounting'],
    createdAt: '2024-01-08T11:30:00Z',
    updatedAt: '2024-01-28T16:45:00Z',
    content: `
RECEIPT

[Company Name]
[Company Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]

RECEIPT DETAILS:
Receipt Number: [Receipt Number]
Date: [Date]
Time: [Time]

RECEIVED FROM:
Name: [Customer Name]
Address: [Customer Address]
[City, State, ZIP Code]
Phone: [Phone Number]
Email: [Email Address]

PAYMENT DETAILS:

Description | Quantity | Unit Price | Total
------------|----------|------------|------
[Item/Service 1] | [Qty] | $[Price] | $[Total]
[Item/Service 2] | [Qty] | $[Price] | $[Total]
[Item/Service 3] | [Qty] | $[Price] | $[Total]

PAYMENT SUMMARY:
Subtotal: $[Subtotal]
Tax ([Tax Rate]%): $[Tax Amount]
Discount: $[Discount]
TOTAL PAID: $[Total Amount]

PAYMENT METHOD:
☐ Cash
☐ Check (Check #: [Check Number])
☐ Credit Card (Last 4 digits: [XXXX])
☐ Bank Transfer
☐ Other: [Specify]

PAYMENT STATUS: PAID IN FULL

ADDITIONAL INFORMATION:
Transaction ID: [Transaction ID]
Reference Number: [Reference Number]

NOTES:
[Any additional notes or terms]

This receipt serves as proof of payment for the above transaction.

Received by: _________________________
[Name], [Title]
[Company Name]

Customer Copy
    `
  },
  {
    id: 'finance-payment-reminder',
    title: 'Payment Reminder Email',
    description: 'Professional email template for following up on overdue payments',
    category: 'finance',
    downloadCount: 4120,
    fileSize: '28.7 KB',
    rating: 4.5,
    tags: ['payment', 'reminder', 'overdue', 'collections', 'email'],
    createdAt: '2024-01-10T14:15:00Z',
    updatedAt: '2024-01-30T10:20:00Z',
    content: `
PAYMENT REMINDER EMAIL

Subject: Payment Reminder - Invoice [Invoice Number] - [Company Name]

Dear [Client Name],

I hope this email finds you well. This is a friendly reminder regarding an outstanding payment on your account.

INVOICE DETAILS:
Invoice Number: [Invoice Number]
Invoice Date: [Invoice Date]
Due Date: [Due Date]
Amount Due: $[Amount Due]
Days Overdue: [Days Overdue]

ORIGINAL INVOICE SUMMARY:
[Brief description of services/products provided]

We have not yet received payment for the above invoice. According to our records, this payment was due on [Due Date].

PAYMENT OPTIONS:
To settle this account, you may:

1. Online Payment: [Payment Portal Link]
2. Bank Transfer:
   Bank Name: [Bank Name]
   Account Number: [Account Number]
   Routing Number: [Routing Number]
3. Check Payment:
   Mail to: [Mailing Address]
   Make payable to: [Company Name]
4. Credit Card: Call [Phone Number]

If you have already sent payment, please disregard this notice. However, if you have any questions about this invoice or need to discuss payment arrangements, please contact me immediately.

NEXT STEPS:
If payment is not received within [Number] days, we may need to:
• Apply late fees as outlined in our terms
• Suspend services until account is current
• Turn the account over to our collections department

We value our business relationship and want to resolve this matter promptly. Please contact me at [Phone Number] or [Email Address] if you have any questions or concerns.

Thank you for your immediate attention to this matter.

Best regards,

[Your Name]
[Your Title]
[Company Name]
[Phone Number]
[Email Address]

ATTACHMENT: Copy of original invoice [Invoice Number]
    `
  },
  {
    id: 'finance-expense-reimbursement',
    title: 'Expense Reimbursement Form',
    description: 'Comprehensive form for employees to request reimbursement of business expenses',
    category: 'finance',
    downloadCount: 2650,
    fileSize: '41.3 KB',
    rating: 4.7,
    tags: ['expense', 'reimbursement', 'employee', 'business', 'travel'],
    createdAt: '2024-01-12T08:45:00Z',
    updatedAt: '2024-01-26T13:15:00Z',
    content: `
EXPENSE REIMBURSEMENT REQUEST FORM

EMPLOYEE INFORMATION:
Name: [Employee Name]
Employee ID: [Employee ID]
Department: [Department]
Manager: [Manager Name]
Date of Request: [Date]

EXPENSE PERIOD:
From: [Start Date]
To: [End Date]

EXPENSE DETAILS:

Date | Category | Description | Business Purpose | Amount | Receipt Attached
-----|----------|-------------|------------------|--------|------------------
[Date] | [Category] | [Description] | [Purpose] | $[Amount] | ☐ Yes ☐ No
[Date] | [Category] | [Description] | [Purpose] | $[Amount] | ☐ Yes ☐ No
[Date] | [Category] | [Description] | [Purpose] | $[Amount] | ☐ Yes ☐ No
[Date] | [Category] | [Description] | [Purpose] | $[Amount] | ☐ Yes ☐ No
[Date] | [Category] | [Description] | [Purpose] | $[Amount] | ☐ Yes ☐ No

EXPENSE CATEGORIES:
☐ Travel (airfare, hotel, car rental)
☐ Meals & Entertainment
☐ Transportation (taxi, parking, mileage)
☐ Office Supplies
☐ Training & Education
☐ Client Entertainment
☐ Communication (phone, internet)
☐ Other: [Specify]

TOTAL AMOUNT REQUESTED: $[Total Amount]

MILEAGE CALCULATION (if applicable):
Starting Location: [Location]
Destination: [Location]
Total Miles: [Miles]
Rate per Mile: $[Rate]
Total Mileage Reimbursement: $[Amount]

PAYMENT INFORMATION:
Preferred Payment Method:
☐ Direct Deposit (use payroll account on file)
☐ Check
☐ Other: [Specify]

If different from payroll account:
Bank Name: [Bank Name]
Account Number: [Account Number]
Routing Number: [Routing Number]

EMPLOYEE CERTIFICATION:
I certify that:
• All expenses listed above were incurred for legitimate business purposes
• I have not been reimbursed for these expenses from any other source
• All receipts and supporting documentation are attached
• The information provided is accurate and complete
• I understand that false claims may result in disciplinary action

Employee Signature: _________________ Date: _______
[Employee Name]

MANAGER APPROVAL:
☐ Approved
☐ Approved with modifications: [Explain]
☐ Denied - Reason: [Reason]

Manager Signature: _________________ Date: _______
[Manager Name]

ACCOUNTING USE ONLY:
Reviewed by: [Accounting Rep]
Date Processed: [Date]
Check Number: [Check Number]
Amount Paid: $[Amount]
GL Account Codes: [Codes]

Notes: [Any special notes or adjustments]

Accounting Signature: _________________ Date: _______
    `
  },
  {
    id: 'finance-budget-planning',
    title: 'Budget Planning Spreadsheet',
    description: 'Comprehensive budget planning template for annual and monthly financial planning',
    category: 'finance',
    downloadCount: 1890,
    fileSize: '67.4 KB',
    rating: 4.9,
    tags: ['budget', 'planning', 'financial', 'annual', 'monthly'],
    createdAt: '2024-01-15T10:30:00Z',
    updatedAt: '2024-01-29T15:45:00Z',
    content: `
ANNUAL BUDGET PLANNING TEMPLATE

COMPANY: [Company Name]
FISCAL YEAR: [Year]
PREPARED BY: [Name]
DATE: [Date]

BUDGET SUMMARY:

REVENUE PROJECTIONS:
Revenue Source | Q1 | Q2 | Q3 | Q4 | Annual Total
---------------|----|----|----|----|-------------
Product Sales | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Service Revenue | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Consulting | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Other Revenue | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
TOTAL REVENUE | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]

EXPENSE PROJECTIONS:

PERSONNEL COSTS:
Expense Category | Q1 | Q2 | Q3 | Q4 | Annual Total
-----------------|----|----|----|----|-------------
Salaries & Wages | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Benefits | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Payroll Taxes | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Training & Development | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
TOTAL PERSONNEL | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]

OPERATING EXPENSES:
Expense Category | Q1 | Q2 | Q3 | Q4 | Annual Total
-----------------|----|----|----|----|-------------
Rent/Lease | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Utilities | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Insurance | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Office Supplies | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Technology | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Marketing | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Travel | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Professional Services | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
TOTAL OPERATING | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]

CAPITAL EXPENDITURES:
Item | Q1 | Q2 | Q3 | Q4 | Annual Total
-----|----|----|----|----|-------------
Equipment | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Software | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Furniture | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Vehicles | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
TOTAL CAPEX | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]

FINANCIAL SUMMARY:
Item | Q1 | Q2 | Q3 | Q4 | Annual Total
-----|----|----|----|----|-------------
Total Revenue | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
Total Expenses | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]
NET INCOME | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Total]

BUDGET ASSUMPTIONS:
• Revenue growth rate: [Percentage]%
• Inflation rate: [Percentage]%
• Employee headcount: [Number] (start) to [Number] (end)
• Major initiatives: [List key projects/initiatives]
• Market conditions: [Brief description]

VARIANCE ANALYSIS (Monthly Tracking):
Month | Budgeted Revenue | Actual Revenue | Variance | Budgeted Expenses | Actual Expenses | Variance
------|------------------|----------------|----------|-------------------|-----------------|----------
Jan | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Amount]
Feb | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Amount]
Mar | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Amount] | $[Amount]

NOTES:
[Additional notes, assumptions, or explanations]

APPROVAL:
Prepared by: _________________ Date: _______
[Name], [Title]

Reviewed by: _________________ Date: _______
[Name], [Title]

Approved by: _________________ Date: _______
[Name], [Title]
    `
  },
  {
    id: 'finance-cash-flow-statement',
    title: 'Cash Flow Statement Template',
    description: 'Professional cash flow statement template for tracking cash inflows and outflows',
    category: 'finance',
    downloadCount: 2340,
    fileSize: '52.8 KB',
    rating: 4.8,
    tags: ['cash-flow', 'financial-statement', 'accounting', 'liquidity', 'operations'],
    createdAt: '2024-01-18T12:00:00Z',
    updatedAt: '2024-01-31T09:30:00Z',
    content: `
CASH FLOW STATEMENT

COMPANY: [Company Name]
PERIOD: [Period - e.g., Year Ended December 31, 2024]
PREPARED BY: [Name]
DATE: [Date]

CASH FLOWS FROM OPERATING ACTIVITIES:

Net Income: $[Amount]

Adjustments to reconcile net income to net cash provided by operating activities:

Non-cash items:
Depreciation and Amortization: $[Amount]
Bad Debt Expense: $[Amount]
Stock-based Compensation: $[Amount]
Gain/Loss on Sale of Assets: $[Amount]
Other Non-cash Items: $[Amount]

Changes in Operating Assets and Liabilities:
Accounts Receivable: $[Amount]
Inventory: $[Amount]
Prepaid Expenses: $[Amount]
Other Current Assets: $[Amount]
Accounts Payable: $[Amount]
Accrued Liabilities: $[Amount]
Deferred Revenue: $[Amount]
Other Current Liabilities: $[Amount]

NET CASH PROVIDED BY (USED IN) OPERATING ACTIVITIES: $[Amount]

CASH FLOWS FROM INVESTING ACTIVITIES:

Capital Expenditures: $([Amount])
Purchase of Equipment: $([Amount])
Purchase of Software: $([Amount])
Sale of Assets: $[Amount]
Investment in Securities: $([Amount])
Sale of Securities: $[Amount]
Acquisition of Business: $([Amount])
Other Investing Activities: $[Amount]

NET CASH PROVIDED BY (USED IN) INVESTING ACTIVITIES: $[Amount]

CASH FLOWS FROM FINANCING ACTIVITIES:

Proceeds from Bank Loans: $[Amount]
Repayment of Bank Loans: $([Amount])
Proceeds from Line of Credit: $[Amount]
Repayment of Line of Credit: $([Amount])
Issuance of Stock: $[Amount]
Repurchase of Stock: $([Amount])
Dividends Paid: $([Amount])
Payment of Lease Obligations: $([Amount])
Other Financing Activities: $[Amount]

NET CASH PROVIDED BY (USED IN) FINANCING ACTIVITIES: $[Amount]

CASH FLOW SUMMARY:
Net Cash from Operating Activities: $[Amount]
Net Cash from Investing Activities: $[Amount]
Net Cash from Financing Activities: $[Amount]

NET INCREASE (DECREASE) IN CASH: $[Amount]

CASH AND CASH EQUIVALENTS:
Beginning of Period: $[Amount]
End of Period: $[Amount]

NET INCREASE (DECREASE) IN CASH: $[Amount]

SUPPLEMENTAL CASH FLOW INFORMATION:
Cash Paid for Interest: $[Amount]
Cash Paid for Income Taxes: $[Amount]

NON-CASH INVESTING AND FINANCING ACTIVITIES:
Equipment Acquired through Capital Lease: $[Amount]
Conversion of Debt to Equity: $[Amount]
Stock Issued for Acquisition: $[Amount]

CASH FLOW RATIOS:
Operating Cash Flow Ratio: [Operating Cash Flow / Current Liabilities]
Cash Coverage Ratio: [Operating Cash Flow / Total Debt]
Cash Flow to Sales Ratio: [Operating Cash Flow / Net Sales]

NOTES:
[Additional explanations or significant events affecting cash flow]

MANAGEMENT ANALYSIS:
[Brief analysis of cash flow trends and implications]

PREPARED BY:
Name: [Name]
Title: [Title]
Date: [Date]
Signature: _________________

REVIEWED BY:
Name: [Name]
Title: [Title]
Date: [Date]
Signature: _________________
    `
  },
  {
    id: 'finance-profit-loss-statement',
    title: 'Profit & Loss Statement',
    description: 'Comprehensive P&L statement template for tracking revenue, expenses, and profitability',
    category: 'finance',
    downloadCount: 3120,
    fileSize: '48.6 KB',
    rating: 4.9,
    tags: ['profit-loss', 'income-statement', 'financial', 'revenue', 'expenses'],
    createdAt: '2024-01-20T14:20:00Z',
    updatedAt: '2024-02-01T11:15:00Z',
    content: `
PROFIT & LOSS STATEMENT
(Income Statement)

COMPANY: [Company Name]
PERIOD: [Period - e.g., Year Ended December 31, 2024]
PREPARED BY: [Name]
DATE: [Date]

REVENUE:
Product Sales: $[Amount]
Service Revenue: $[Amount]
Consulting Revenue: $[Amount]
Licensing Revenue: $[Amount]
Other Revenue: $[Amount]
TOTAL REVENUE: $[Amount]

COST OF GOODS SOLD (COGS):
Direct Materials: $[Amount]
Direct Labor: $[Amount]
Manufacturing Overhead: $[Amount]
Cost of Services: $[Amount]
Other Direct Costs: $[Amount]
TOTAL COST OF GOODS SOLD: $[Amount]

GROSS PROFIT: $[Amount]
Gross Profit Margin: [Percentage]%

OPERATING EXPENSES:

Sales & Marketing:
Advertising: $[Amount]
Marketing Campaigns: $[Amount]
Sales Commissions: $[Amount]
Trade Shows: $[Amount]
Sales Travel: $[Amount]
Total Sales & Marketing: $[Amount]

General & Administrative:
Salaries & Wages: $[Amount]
Employee Benefits: $[Amount]
Payroll Taxes: $[Amount]
Rent: $[Amount]
Utilities: $[Amount]
Insurance: $[Amount]
Office Supplies: $[Amount]
Professional Services: $[Amount]
Depreciation: $[Amount]
Other G&A: $[Amount]
Total General & Administrative: $[Amount]

Research & Development:
R&D Salaries: $[Amount]
R&D Materials: $[Amount]
R&D Equipment: $[Amount]
Other R&D: $[Amount]
Total Research & Development: $[Amount]

TOTAL OPERATING EXPENSES: $[Amount]

OPERATING INCOME (EBITDA): $[Amount]
Operating Margin: [Percentage]%

OTHER INCOME (EXPENSE):
Interest Income: $[Amount]
Interest Expense: $([Amount])
Investment Gains/Losses: $[Amount]
Foreign Exchange Gains/Losses: $[Amount]
Other Income: $[Amount]
TOTAL OTHER INCOME (EXPENSE): $[Amount]

INCOME BEFORE TAXES: $[Amount]

INCOME TAX EXPENSE:
Federal Income Tax: $[Amount]
State Income Tax: $[Amount]
Other Taxes: $[Amount]
TOTAL INCOME TAX EXPENSE: $[Amount]

NET INCOME: $[Amount]
Net Profit Margin: [Percentage]%

EARNINGS PER SHARE (if applicable):
Basic EPS: $[Amount]
Diluted EPS: $[Amount]
Weighted Average Shares Outstanding: [Number]

KEY FINANCIAL RATIOS:
Gross Profit Margin: [Percentage]%
Operating Margin: [Percentage]%
Net Profit Margin: [Percentage]%
Return on Assets (ROA): [Percentage]%
Return on Equity (ROE): [Percentage]%

COMPARATIVE ANALYSIS:
Item | Current Period | Prior Period | Change | % Change
-----|----------------|--------------|--------|----------
Total Revenue | $[Amount] | $[Amount] | $[Amount] | [%]
Gross Profit | $[Amount] | $[Amount] | $[Amount] | [%]
Operating Income | $[Amount] | $[Amount] | $[Amount] | [%]
Net Income | $[Amount] | $[Amount] | $[Amount] | [%]

NOTES TO FINANCIAL STATEMENTS:
[Significant accounting policies, events, or explanations]

MANAGEMENT COMMENTARY:
[Brief analysis of financial performance and key drivers]

PREPARED BY:
Name: [Name]
Title: [Title]
Date: [Date]
Signature: _________________

REVIEWED BY:
Name: [Name]
Title: [Title]
Date: [Date]
Signature: _________________
    `
  },
  {
    id: 'finance-balance-sheet',
    title: 'Balance Sheet',
    description: 'Professional balance sheet template showing assets, liabilities, and equity',
    category: 'finance',
    downloadCount: 2780,
    fileSize: '44.9 KB',
    rating: 4.8,
    tags: ['balance-sheet', 'assets', 'liabilities', 'equity', 'financial-position'],
    createdAt: '2024-01-22T16:45:00Z',
    updatedAt: '2024-02-02T13:20:00Z',
    content: `
BALANCE SHEET

COMPANY: [Company Name]
AS OF: [Date - e.g., December 31, 2024]
PREPARED BY: [Name]
DATE: [Date]

ASSETS

CURRENT ASSETS:
Cash and Cash Equivalents: $[Amount]
Short-term Investments: $[Amount]
Accounts Receivable: $[Amount]
Less: Allowance for Doubtful Accounts: $([Amount])
Net Accounts Receivable: $[Amount]
Inventory: $[Amount]
Prepaid Expenses: $[Amount]
Other Current Assets: $[Amount]
TOTAL CURRENT ASSETS: $[Amount]

NON-CURRENT ASSETS:

Property, Plant & Equipment:
Land: $[Amount]
Buildings: $[Amount]
Equipment: $[Amount]
Furniture & Fixtures: $[Amount]
Vehicles: $[Amount]
Total Property, Plant & Equipment: $[Amount]
Less: Accumulated Depreciation: $([Amount])
Net Property, Plant & Equipment: $[Amount]

Intangible Assets:
Goodwill: $[Amount]
Patents: $[Amount]
Trademarks: $[Amount]
Software: $[Amount]
Other Intangible Assets: $[Amount]
Total Intangible Assets: $[Amount]
Less: Accumulated Amortization: $([Amount])
Net Intangible Assets: $[Amount]

Other Non-Current Assets:
Long-term Investments: $[Amount]
Deferred Tax Assets: $[Amount]
Other Assets: $[Amount]
TOTAL OTHER NON-CURRENT ASSETS: $[Amount]

TOTAL NON-CURRENT ASSETS: $[Amount]

TOTAL ASSETS: $[Amount]

LIABILITIES AND STOCKHOLDERS' EQUITY

CURRENT LIABILITIES:
Accounts Payable: $[Amount]
Accrued Liabilities: $[Amount]
Short-term Debt: $[Amount]
Current Portion of Long-term Debt: $[Amount]
Deferred Revenue: $[Amount]
Income Taxes Payable: $[Amount]
Other Current Liabilities: $[Amount]
TOTAL CURRENT LIABILITIES: $[Amount]

NON-CURRENT LIABILITIES:
Long-term Debt: $[Amount]
Deferred Tax Liabilities: $[Amount]
Pension Obligations: $[Amount]
Other Long-term Liabilities: $[Amount]
TOTAL NON-CURRENT LIABILITIES: $[Amount]

TOTAL LIABILITIES: $[Amount]

STOCKHOLDERS' EQUITY:
Common Stock ([Number] shares authorized, [Number] shares issued): $[Amount]
Preferred Stock ([Number] shares authorized, [Number] shares issued): $[Amount]
Additional Paid-in Capital: $[Amount]
Retained Earnings: $[Amount]
Accumulated Other Comprehensive Income (Loss): $[Amount]
Treasury Stock ([Number] shares): $([Amount])
TOTAL STOCKHOLDERS' EQUITY: $[Amount]

TOTAL LIABILITIES AND STOCKHOLDERS' EQUITY: $[Amount]

FINANCIAL RATIOS:

Liquidity Ratios:
Current Ratio: [Current Assets / Current Liabilities]
Quick Ratio: [(Current Assets - Inventory) / Current Liabilities]
Cash Ratio: [Cash / Current Liabilities]

Leverage Ratios:
Debt-to-Equity Ratio: [Total Debt / Total Equity]
Debt-to-Assets Ratio: [Total Debt / Total Assets]
Equity Ratio: [Total Equity / Total Assets]

Efficiency Ratios:
Asset Turnover: [Revenue / Total Assets]
Inventory Turnover: [COGS / Average Inventory]
Receivables Turnover: [Revenue / Average Accounts Receivable]

COMPARATIVE BALANCE SHEET:
Item | Current Year | Prior Year | Change | % Change
-----|--------------|------------|--------|----------
Total Assets | $[Amount] | $[Amount] | $[Amount] | [%]
Total Liabilities | $[Amount] | $[Amount] | $[Amount] | [%]
Total Equity | $[Amount] | $[Amount] | $[Amount] | [%]

NOTES TO BALANCE SHEET:
[Significant accounting policies, commitments, contingencies, or explanations]

SUBSEQUENT EVENTS:
[Any significant events occurring after the balance sheet date]

PREPARED BY:
Name: [Name]
Title: [Title]
Date: [Date]
Signature: _________________

REVIEWED BY:
Name: [Name]
Title: [Title]
Date: [Date]
Signature: _________________
    `
  },
  {
    id: 'finance-purchase-order',
    title: 'Purchase Order Template',
    description: 'Professional purchase order template for ordering goods and services from vendors',
    category: 'finance',
    downloadCount: 3890,
    fileSize: '38.4 KB',
    rating: 4.6,
    tags: ['purchase-order', 'procurement', 'vendor', 'ordering', 'supplies'],
    createdAt: '2024-01-25T09:30:00Z',
    updatedAt: '2024-02-03T14:45:00Z',
    content: `
PURCHASE ORDER

[Company Name]
[Company Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]

VENDOR INFORMATION:
Vendor Name: [Vendor Name]
Contact Person: [Contact Name]
Address: [Vendor Address]
[City, State, ZIP Code]
Phone: [Phone Number]
Email: [Email Address]
Vendor ID: [Vendor ID]

PURCHASE ORDER DETAILS:
PO Number: [PO Number]
PO Date: [Date]
Required Delivery Date: [Date]
Requested by: [Requestor Name]
Department: [Department]
Project/Job Number: [Project Number]

SHIPPING INFORMATION:
Ship To:
[Company/Department Name]
[Shipping Address]
[City, State, ZIP Code]

Special Shipping Instructions:
[Instructions]

BILLING INFORMATION:
Bill To:
[Company Name]
[Billing Address]
[City, State, ZIP Code]

ITEM DETAILS:

Item # | Description | Quantity | Unit | Unit Price | Total Price
-------|-------------|----------|------|------------|------------
[Item 1] | [Description] | [Qty] | [Unit] | $[Price] | $[Total]
[Item 2] | [Description] | [Qty] | [Unit] | $[Price] | $[Total]
[Item 3] | [Description] | [Qty] | [Unit] | $[Price] | $[Total]
[Item 4] | [Description] | [Qty] | [Unit] | $[Price] | $[Total]
[Item 5] | [Description] | [Qty] | [Unit] | $[Price] | $[Total]

COST SUMMARY:
Subtotal: $[Subtotal]
Shipping & Handling: $[Shipping]
Tax: $[Tax Amount]
Other Charges: $[Other]
TOTAL AMOUNT: $[Total Amount]

PAYMENT TERMS:
☐ Net 30
☐ Net 15
☐ Due on Receipt
☐ 2/10 Net 30
☐ Other: [Specify]

DELIVERY TERMS:
☐ FOB Shipping Point
☐ FOB Destination
☐ Other: [Specify]

SPECIAL TERMS AND CONDITIONS:
• All items must be delivered by the required delivery date
• Partial shipments are [acceptable/not acceptable]
• All items must meet specified quality standards
• Vendor must provide advance notice of any delays
• Returns require prior authorization
• [Additional terms]

APPROVAL:
Requested by: _________________ Date: _______
[Name], [Title]

Approved by: _________________ Date: _______
[Name], [Title]

Budget Approval: _________________ Date: _______
[Name], [Title]

VENDOR ACKNOWLEDGMENT:
By signing below, vendor acknowledges receipt of this purchase order and agrees to the terms and conditions stated above.

Vendor Signature: _________________ Date: _______
[Name], [Title]
[Company Name]

RECEIVING INFORMATION (To be completed upon receipt):
Date Received: [Date]
Received by: [Name]
Quantity Received: [Quantity]
Condition: ☐ Good ☐ Damaged ☐ Incomplete
Comments: [Comments]

Receiver Signature: _________________ Date: _______
    `
  },
  {
    id: 'finance-bill-of-sale',
    title: 'Bill of Sale',
    description: 'Legal bill of sale template for documenting the transfer of ownership of goods',
    category: 'finance',
    downloadCount: 2450,
    fileSize: '35.7 KB',
    rating: 4.5,
    tags: ['bill-of-sale', 'ownership', 'transfer', 'legal', 'transaction'],
    createdAt: '2024-01-28T11:15:00Z',
    updatedAt: '2024-02-04T16:30:00Z',
    content: `
BILL OF SALE

SELLER INFORMATION:
Name: [Seller Name]
Address: [Seller Address]
[City, State, ZIP Code]
Phone: [Phone Number]
Email: [Email Address]
Driver's License #: [License Number]

BUYER INFORMATION:
Name: [Buyer Name]
Address: [Buyer Address]
[City, State, ZIP Code]
Phone: [Phone Number]
Email: [Email Address]
Driver's License #: [License Number]

TRANSACTION DETAILS:
Date of Sale: [Date]
Location of Sale: [Location]

ITEM(S) BEING SOLD:

Description: [Detailed description of item(s)]
Make: [Make]
Model: [Model]
Year: [Year]
Serial Number/VIN: [Serial Number]
Color: [Color]
Condition: [Condition]
Mileage (if applicable): [Mileage]

Additional Items Included:
☐ [Item 1]
☐ [Item 2]
☐ [Item 3]
☐ Other: [Specify]

FINANCIAL TERMS:
Purchase Price: $[Amount]

Payment Method:
☐ Cash
☐ Cashier's Check
☐ Personal Check (Check #: [Number])
☐ Bank Transfer
☐ Other: [Specify]

Payment Schedule:
☐ Paid in Full
☐ Down Payment: $[Amount] (Balance: $[Amount] due by [Date])
☐ Installment Plan: [Details]

CONDITION AND WARRANTIES:
The item is sold in "AS IS" condition with the following warranties:

☐ No warranties expressed or implied
☐ Limited warranty: [Specify terms]
☐ Full warranty for [Time period]

Known Defects or Issues:
[List any known problems or defects]

SELLER'S REPRESENTATIONS:
The Seller represents and warrants that:
• Seller is the legal owner of the item(s) being sold
• The item is free and clear of all liens and encumbrances
• Seller has the right to sell the item
• All information provided is true and accurate
• [Additional representations]

BUYER'S ACKNOWLEDGMENTS:
The Buyer acknowledges that:
• Buyer has inspected the item and accepts its condition
• Buyer understands the terms of this sale
• Buyer is purchasing the item "as is" unless otherwise stated
• Buyer is responsible for all applicable taxes and fees
• [Additional acknowledgments]

ADDITIONAL TERMS:
• Title transfer: [Details about title transfer process]
• Registration: Buyer is responsible for registration and licensing
• Insurance: [Insurance requirements if applicable]
• Delivery: [Delivery arrangements]
• [Other terms and conditions]

SIGNATURES:

SELLER:
I hereby certify that I am the legal owner of the above-described item(s) and have the right to sell them. I transfer all ownership rights to the Buyer.

Seller Signature: _________________ Date: _______
Print Name: [Seller Name]

BUYER:
I acknowledge that I have read and understand this Bill of Sale and agree to purchase the above-described item(s) under the terms stated.

Buyer Signature: _________________ Date: _______
Print Name: [Buyer Name]

WITNESS (if required):
Witness Signature: _________________ Date: _______
Print Name: [Witness Name]
Address: [Witness Address]

NOTARIZATION (if required):
State of [State]
County of [County]

On this [Date], before me personally appeared [Names], who proved to me on the basis of satisfactory evidence to be the persons whose names are subscribed to the within instrument and acknowledged to me that they executed the same in their authorized capacities.

Notary Public: _________________
My Commission Expires: [Date]
    `
  },
  {
    id: 'finance-tax-deduction-tracker',
    title: 'Tax Deduction Tracker',
    description: 'Comprehensive tracker for organizing and documenting tax-deductible business expenses',
    category: 'finance',
    downloadCount: 3670,
    fileSize: '49.3 KB',
    rating: 4.7,
    tags: ['tax', 'deductions', 'expenses', 'irs', 'business'],
    createdAt: '2024-01-30T13:45:00Z',
    updatedAt: '2024-02-05T10:20:00Z',
    content: `
TAX DEDUCTION TRACKER

TAXPAYER INFORMATION:
Name/Business: [Name]
Tax Year: [Year]
EIN/SSN: [Number]
Business Type: [Sole Proprietorship/LLC/Corporation/Partnership]
Prepared by: [Name]
Date: [Date]

BUSINESS EXPENSE CATEGORIES:

OFFICE EXPENSES:
Date | Description | Amount | Receipt # | Notes
-----|-------------|--------|-----------|-------
[Date] | [Description] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Description] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Description] | $[Amount] | [Receipt #] | [Notes]
TOTAL OFFICE EXPENSES: $[Total]

TRAVEL & TRANSPORTATION:
Date | Destination | Purpose | Mileage | Amount | Receipt # | Notes
-----|-------------|---------|---------|--------|-----------|-------
[Date] | [Destination] | [Purpose] | [Miles] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Destination] | [Purpose] | [Miles] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Destination] | [Purpose] | [Miles] | $[Amount] | [Receipt #] | [Notes]
TOTAL TRAVEL EXPENSES: $[Total]
TOTAL MILEAGE: [Miles] × $[Rate per mile] = $[Total]

MEALS & ENTERTAINMENT:
Date | Description | Business Purpose | Attendees | Amount | Receipt # | Deductible %
-----|-------------|------------------|-----------|--------|-----------|-------------
[Date] | [Description] | [Purpose] | [Attendees] | $[Amount] | [Receipt #] | [%]
[Date] | [Description] | [Purpose] | [Attendees] | $[Amount] | [Receipt #] | [%]
[Date] | [Description] | [Purpose] | [Attendees] | $[Amount] | [Receipt #] | [%]
TOTAL MEALS & ENTERTAINMENT: $[Total]

PROFESSIONAL SERVICES:
Date | Service Provider | Service Type | Amount | Receipt # | Notes
-----|------------------|--------------|--------|-----------|-------
[Date] | [Provider] | [Service] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Provider] | [Service] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Provider] | [Service] | $[Amount] | [Receipt #] | [Notes]
TOTAL PROFESSIONAL SERVICES: $[Total]

EQUIPMENT & SUPPLIES:
Date | Item | Business Use % | Cost | Depreciation | Deductible Amount | Receipt #
-----|------|----------------|------|--------------|-------------------|----------
[Date] | [Item] | [%] | $[Cost] | $[Depreciation] | $[Amount] | [Receipt #]
[Date] | [Item] | [%] | $[Cost] | $[Depreciation] | $[Amount] | [Receipt #]
[Date] | [Item] | [%] | $[Cost] | $[Depreciation] | $[Amount] | [Receipt #]
TOTAL EQUIPMENT & SUPPLIES: $[Total]

HOME OFFICE EXPENSES:
Total Home Square Footage: [Sq Ft]
Office Square Footage: [Sq Ft]
Business Use Percentage: [%]

Expense Type | Annual Amount | Business Portion | Deductible Amount
-------------|---------------|------------------|------------------
Mortgage Interest/Rent | $[Amount] | [%] | $[Amount]
Property Taxes | $[Amount] | [%] | $[Amount]
Utilities | $[Amount] | [%] | $[Amount]
Insurance | $[Amount] | [%] | $[Amount]
Repairs & Maintenance | $[Amount] | [%] | $[Amount]
Depreciation | $[Amount] | [%] | $[Amount]
TOTAL HOME OFFICE: $[Total]

MARKETING & ADVERTISING:
Date | Description | Type | Amount | Receipt # | Notes
-----|-------------|------|--------|-----------|-------
[Date] | [Description] | [Type] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Description] | [Type] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Description] | [Type] | $[Amount] | [Receipt #] | [Notes]
TOTAL MARKETING & ADVERTISING: $[Total]

INSURANCE PREMIUMS:
Insurance Type | Provider | Annual Premium | Business Portion | Deductible Amount
---------------|----------|----------------|------------------|------------------
[Type] | [Provider] | $[Amount] | [%] | $[Amount]
[Type] | [Provider] | $[Amount] | [%] | $[Amount]
[Type] | [Provider] | $[Amount] | [%] | $[Amount]
TOTAL INSURANCE: $[Total]

EDUCATION & TRAINING:
Date | Course/Event | Provider | Amount | Receipt # | Business Relevance
-----|--------------|----------|--------|-----------|-------------------
[Date] | [Course] | [Provider] | $[Amount] | [Receipt #] | [Relevance]
[Date] | [Course] | [Provider] | $[Amount] | [Receipt #] | [Relevance]
[Date] | [Course] | [Provider] | $[Amount] | [Receipt #] | [Relevance]
TOTAL EDUCATION & TRAINING: $[Total]

OTHER BUSINESS EXPENSES:
Date | Description | Category | Amount | Receipt # | Notes
-----|-------------|----------|--------|-----------|-------
[Date] | [Description] | [Category] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Description] | [Category] | $[Amount] | [Receipt #] | [Notes]
[Date] | [Description] | [Category] | $[Amount] | [Receipt #] | [Notes]
TOTAL OTHER EXPENSES: $[Total]

SUMMARY OF DEDUCTIONS:
Office Expenses: $[Amount]
Travel & Transportation: $[Amount]
Meals & Entertainment: $[Amount]
Professional Services: $[Amount]
Equipment & Supplies: $[Amount]
Home Office: $[Amount]
Marketing & Advertising: $[Amount]
Insurance: $[Amount]
Education & Training: $[Amount]
Other Expenses: $[Amount]

TOTAL BUSINESS DEDUCTIONS: $[Total Amount]

DOCUMENTATION CHECKLIST:
☐ All receipts organized and filed
☐ Mileage log completed
☐ Home office measurements documented
☐ Business purpose documented for all expenses
☐ Credit card statements reconciled
☐ Bank statements reviewed
☐ Depreciation schedules updated

NOTES:
[Additional notes, explanations, or reminders for tax preparation]

PREPARED BY:
Name: [Name]
Date: [Date]
Signature: _________________

REVIEWED BY TAX PROFESSIONAL:
Name: [Tax Professional Name]
Date: [Date]
Signature: _________________
    `
  }
];

// Add finance templates to the main templates array
templates.push(...financeTemplates);

// Sales Templates
const salesTemplates: DocumentTemplate[] = [
  {
    id: 'sales-quotation-template',
    title: 'Sales Quotation Template',
    description: 'Professional quotation template for providing detailed pricing and terms to potential clients',
    category: 'sales',
    downloadCount: 4120,
    fileSize: '42.3 KB',
    rating: 4.8,
    tags: ['quotation', 'quote', 'pricing', 'proposal', 'sales'],
    createdAt: '2024-01-06T10:00:00Z',
    updatedAt: '2024-01-27T14:30:00Z',
    content: `
SALES QUOTATION

[Company Name]
[Company Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]
[Website]

QUOTATION TO:
[Client Name]
[Company Name]
[Client Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]

QUOTATION DETAILS:
Quote Number: [Quote Number]
Date: [Date]
Valid Until: [Expiration Date]
Prepared By: [Sales Rep Name]
Sales Representative: [Sales Rep Name]
Phone: [Sales Rep Phone]
Email: [Sales Rep Email]

PROJECT INFORMATION:
Project Name: [Project Name]
Project Description: [Brief description of project/requirements]
Estimated Start Date: [Date]
Estimated Completion Date: [Date]

ITEMIZED QUOTE:

Item | Description | Quantity | Unit Price | Total
-----|-------------|----------|------------|-------
[Item 1] | [Detailed description] | [Qty] | $[Price] | $[Total]
[Item 2] | [Detailed description] | [Qty] | $[Price] | $[Total]
[Item 3] | [Detailed description] | [Qty] | $[Price] | $[Total]
[Item 4] | [Detailed description] | [Qty] | $[Price] | $[Total]
[Item 5] | [Detailed description] | [Qty] | $[Price] | $[Total]

PRICING SUMMARY:
Subtotal: $[Subtotal]
Discount ([Percentage]%): $[Discount]
Tax ([Tax Rate]%): $[Tax Amount]
Shipping & Handling: $[Shipping]
TOTAL QUOTED PRICE: $[Total Amount]

PAYMENT TERMS:
• Payment Schedule: [e.g., 50% upfront, 50% upon completion]
• Payment Methods Accepted: [Check, Wire Transfer, Credit Card, etc.]
• Late Payment Fee: [Percentage]% per month
• Currency: [Currency Type]

DELIVERY & IMPLEMENTATION:
• Estimated Delivery Time: [Timeframe]
• Delivery Location: [Location]
• Installation/Setup: [Included/Not Included/Additional Cost]
• Training: [Included/Not Included/Additional Cost]

TERMS & CONDITIONS:
• This quotation is valid for [Number] days from the date above
• Prices are subject to change after expiration date
• All work/products subject to availability
• Minimum order requirements: [If applicable]
• Cancellation policy: [Policy details]
• Warranty: [Warranty terms]
• Returns: [Return policy]

EXCLUSIONS:
The following items are NOT included in this quote:
• [Exclusion 1]
• [Exclusion 2]
• [Exclusion 3]

ASSUMPTIONS:
This quote is based on the following assumptions:
• [Assumption 1]
• [Assumption 2]
• [Assumption 3]

NEXT STEPS:
To accept this quotation:
1. Review all terms and conditions
2. Sign and date the acceptance section below
3. Submit payment as per payment terms
4. Return signed copy to [Email/Address]

We look forward to working with you on this project. Please contact us if you have any questions.

ACCEPTANCE:
By signing below, I accept the terms and conditions outlined in this quotation.

Client Signature: _________________ Date: _______
Print Name: [Client Name]
Title: [Client Title]

[Your Name]
[Your Title]
[Company Name]
    `
  },
  {
    id: 'sales-contract-agreement',
    title: 'Sales Contract Agreement',
    description: 'Comprehensive sales contract with legal terms for goods or services transactions',
    category: 'sales',
    downloadCount: 3780,
    fileSize: '58.9 KB',
    rating: 4.9,
    tags: ['contract', 'agreement', 'legal', 'terms', 'sales'],
    createdAt: '2024-01-08T09:30:00Z',
    updatedAt: '2024-01-29T16:45:00Z',
    content: `
SALES CONTRACT AGREEMENT

This Sales Contract Agreement ("Agreement") is entered into on [Date] ("Effective Date") by and between:

SELLER:
[Company Name] ("Seller")
[Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]

BUYER:
[Company Name] ("Buyer")
[Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]

1. PRODUCTS/SERVICES
The Seller agrees to sell, and the Buyer agrees to purchase, the following products/services:

Product/Service | Description | Quantity | Unit Price | Total
----------------|-------------|----------|------------|-------
[Item 1] | [Description] | [Qty] | $[Price] | $[Total]
[Item 2] | [Description] | [Qty] | $[Price] | $[Total]
[Item 3] | [Description] | [Qty] | $[Price] | $[Total]

Total Contract Value: $[Total Amount]

2. PRICE AND PAYMENT TERMS
2.1 Total Purchase Price: $[Amount]
2.2 Payment Schedule:
    • Deposit: $[Amount] due upon signing
    • Progress Payment(s): $[Amount] due [Milestone]
    • Final Payment: $[Amount] due [Condition]
2.3 Payment Method: [Specify accepted methods]
2.4 Late Payment: Interest at [Rate]% per month on overdue amounts

3. DELIVERY TERMS
3.1 Delivery Date: [Date or timeframe]
3.2 Delivery Location: [Address]
3.3 Shipping Method: [Method]
3.4 Risk of Loss: Passes to Buyer upon [Delivery/Shipment]
3.5 Shipping Costs: [Seller pays/Buyer pays/Split]

4. INSPECTION AND ACCEPTANCE
4.1 Buyer has [Number] days from delivery to inspect products
4.2 Buyer must notify Seller of any defects within inspection period
4.3 Acceptance deemed complete if no notice provided within inspection period
4.4 Rejected products must be returned within [Number] days

5. WARRANTIES
5.1 Seller warrants that:
    • Products are free from defects in materials and workmanship
    • Products conform to specifications
    • Seller has clear title to products
    • Products are free from liens and encumbrances
5.2 Warranty Period: [Duration]
5.3 Warranty excludes: [List exclusions]

6. LIMITATION OF LIABILITY
6.1 Seller's total liability shall not exceed the contract price
6.2 Neither party liable for indirect, consequential, or punitive damages
6.3 Exceptions: [List any exceptions]

7. INTELLECTUAL PROPERTY
7.1 All intellectual property rights remain with [Seller/Buyer]
7.2 License granted: [Describe any licenses]
7.3 Restrictions: [List any restrictions]

8. CONFIDENTIALITY
8.1 Both parties agree to maintain confidentiality of proprietary information
8.2 Confidentiality obligations survive termination for [Duration]
8.3 Exceptions: Publicly available information, independently developed

9. FORCE MAJEURE
Neither party liable for delays due to circumstances beyond reasonable control, including natural disasters, war, labor disputes, or government actions.

10. TERMINATION
10.1 Either party may terminate for material breach with [Notice Period] notice
10.2 Buyer may terminate for convenience with [Notice Period] notice and payment of [Terms]
10.3 Upon termination, Buyer shall pay for products delivered and accepted

11. DISPUTE RESOLUTION
11.1 Parties agree to first attempt resolution through good faith negotiation
11.2 If unresolved, disputes shall be settled by [Mediation/Arbitration]
11.3 Venue: [Location]
11.4 Governing Law: [State/Country]

12. INSURANCE
[Specify insurance requirements for both parties]

13. COMPLIANCE
Both parties shall comply with all applicable laws, regulations, and industry standards.

14. AMENDMENTS
This Agreement may only be amended by written agreement signed by both parties.

15. ASSIGNMENT
Neither party may assign this Agreement without prior written consent of the other party.

16. ENTIRE AGREEMENT
This Agreement constitutes the entire agreement and supersedes all prior negotiations, representations, or agreements.

17. SEVERABILITY
If any provision is found invalid, the remaining provisions shall remain in full force.

18. NOTICES
All notices shall be in writing and delivered to the addresses specified above.

19. COUNTERPARTS
This Agreement may be executed in counterparts, each constituting an original.

SELLER:                           BUYER:

_________________________         _________________________
[Name], [Title]                   [Name], [Title]
[Company Name]                    [Company Name]
Date: _______________             Date: _______________

WITNESS (if required):
_________________________
[Name]
Date: _______________
    `
  },
  {
    id: 'sales-lead-tracking',
    title: 'Lead Tracking Sheet',
    description: 'Comprehensive tracking system for managing sales leads through the pipeline',
    category: 'sales',
    downloadCount: 5230,
    fileSize: '51.4 KB',
    rating: 4.7,
    tags: ['leads', 'tracking', 'pipeline', 'crm', 'sales-management'],
    createdAt: '2024-01-10T11:15:00Z',
    updatedAt: '2024-01-30T13:20:00Z',
    content: `
LEAD TRACKING SHEET

COMPANY: [Company Name]
PERIOD: [Month/Quarter/Year]
SALES REP: [Sales Representative Name]
LAST UPDATED: [Date]

LEAD INFORMATION TRACKER:

Lead ID | Date Added | Source | Lead Name | Company | Contact Info | Industry | Status | Priority | Value | Stage | Next Action | Follow-Up Date | Notes
--------|------------|--------|-----------|---------|--------------|----------|--------|----------|-------|-------|-------------|----------------|-------
[ID] | [Date] | [Source] | [Name] | [Company] | [Email/Phone] | [Industry] | [Status] | [Priority] | $[Value] | [Stage] | [Action] | [Date] | [Notes]
[ID] | [Date] | [Source] | [Name] | [Company] | [Email/Phone] | [Industry] | [Status] | [Priority] | $[Value] | [Stage] | [Action] | [Date] | [Notes]
[ID] | [Date] | [Source] | [Name] | [Company] | [Email/Phone] | Industry] | [Status] | [Priority] | $[Value] | [Stage] | [Action] | [Date] | [Notes]

LEAD SOURCE CATEGORIES:
• Website Inquiry
• Trade Show
• Referral
• Cold Call
• Email Campaign
• Social Media
• Partner
• Event
• Advertisement
• Other

LEAD STATUS:
• New - Just received, not yet contacted
• Contacted - Initial contact made
• Qualified - Meets buying criteria
• Proposal Sent - Quote/proposal provided
• Negotiating - Discussing terms
• Won - Converted to customer
• Lost - Did not convert
• Nurturing - Not ready to buy yet

PRIORITY LEVELS:
• Hot - High probability, immediate action
• Warm - Interested, follow up soon
• Cold - Low interest, periodic follow-up

SALES PIPELINE STAGES:
1. Lead Generation
2. Initial Contact
3. Qualification
4. Needs Analysis
5. Proposal/Quote
6. Negotiation
7. Closing
8. Won/Lost

DETAILED LEAD RECORDS:

LEAD #1:
Lead ID: [ID]
Date Added: [Date]
Lead Source: [Source]
Referral Source (if applicable): [Name/Company]

Contact Information:
Name: [Full Name]
Title: [Job Title]
Company: [Company Name]
Email: [Email]
Phone: [Phone Number]
LinkedIn: [URL]
Address: [Address]

Company Information:
Industry: [Industry]
Company Size: [Number of employees]
Annual Revenue: [Revenue range]
Website: [URL]
Decision Maker: [Yes/No]
Budget Authority: [Yes/No]

Lead Qualification:
Need: [Describe need]
Budget: $[Budget range]
Timeline: [When looking to buy]
Authority: [Decision maker contact]
Qualification Score: [Score/10]

Current Status: [Status]
Pipeline Stage: [Stage]
Priority: [Hot/Warm/Cold]
Estimated Deal Value: $[Amount]
Probability of Close: [Percentage]%
Estimated Close Date: [Date]

Activity History:
Date | Activity Type | Notes | Next Step
-----|---------------|-------|----------
[Date] | [Call/Email/Meeting] | [Details] | [Next action]
[Date] | [Call/Email/Meeting] | [Details] | [Next action]
[Date] | [Call/Email/Meeting] | [Details] | [Next action]

Next Action Required:
Action: [Specific action]
Due Date: [Date]
Assigned To: [Sales Rep]

Objections/Concerns:
• [Objection 1]
• [Objection 2]
• [Objection 3]

Competition:
Current Provider: [Competitor name]
Also Considering: [Other competitors]
Our Advantages: [List advantages]

PIPELINE METRICS SUMMARY:

Stage | Number of Leads | Total Value | Average Value | Conversion Rate
------|-----------------|-------------|---------------|----------------
Lead Generation | [Number] | $[Total] | $[Average] | [%]
Initial Contact | [Number] | $[Total] | $[Average] | [%]
Qualification | [Number] | $[Total] | $[Average] | [%]
Needs Analysis | [Number] | $[Total] | $[Average] | [%]
Proposal/Quote | [Number] | $[Total] | $[Average] | [%]
Negotiation | [Number] | $[Total] | $[Average] | [%]
Closing | [Number] | $[Total] | $[Average] | [%]

PERFORMANCE METRICS:

Total Leads: [Number]
New Leads This Period: [Number]
Active Leads: [Number]
Qualified Leads: [Number]
Proposals Sent: [Number]
Deals Won: [Number]
Deals Lost: [Number]
Win Rate: [Percentage]%
Total Pipeline Value: $[Amount]
Average Deal Size: $[Amount]
Average Sales Cycle: [Days]

LEAD SOURCE PERFORMANCE:

Source | Leads | Qualified | Won | Win Rate | Total Value
-------|-------|-----------|-----|----------|------------
[Source 1] | [#] | [#] | [#] | [%] | $[Amount]
[Source 2] | [#] | [#] | [#] | [%] | $[Amount]
[Source 3] | [#] | [#] | [#] | [%] | $[Amount]

FOLLOW-UP SCHEDULE:

This Week:
Lead Name | Company | Action Required | Due Date | Status
----------|---------|-----------------|----------|--------
[Name] | [Company] | [Action] | [Date] | [Status]
[Name] | [Company] | [Action] | [Date] | [Status]

Next Week:
Lead Name | Company | Action Required | Due Date | Status
----------|---------|-----------------|----------|--------
[Name] | [Company] | [Action] | [Date] | [Status]
[Name] | [Company] | [Action] | [Date] | [Status]

NOTES & OBSERVATIONS:
[Record trends, insights, and strategies]

Prepared by: _________________ Date: _______
[Sales Manager Name]
    `
  },
  {
    id: 'sales-follow-up-email',
    title: 'Sales Follow-Up Email Sequence',
    description: 'Series of professional follow-up email templates for different stages of the sales process',
    category: 'sales',
    downloadCount: 6450,
    fileSize: '39.7 KB',
    rating: 4.8,
    tags: ['follow-up', 'email', 'sequence', 'outreach', 'communication'],
    createdAt: '2024-01-12T14:30:00Z',
    updatedAt: '2024-02-01T10:15:00Z',
    content: `
SALES FOLLOW-UP EMAIL SEQUENCE

EMAIL #1: INITIAL FOLLOW-UP (Day 1 after meeting/call)

Subject: Great speaking with you, [First Name]!

Hi [First Name],

Thank you for taking the time to speak with me [today/yesterday] about [specific topic discussed]. I enjoyed learning more about [Company Name] and your [specific challenge/goal].

KEY TAKEAWAYS FROM OUR CONVERSATION:
• [Point 1 from discussion]
• [Point 2 from discussion]
• [Point 3 from discussion]

NEXT STEPS:
Based on our discussion, I'm attaching [resource/information] that addresses [specific need]. I believe this will help you [specific benefit].

I'd like to schedule a follow-up call on [date/time options] to [next step purpose]. Does this work with your schedule?

Looking forward to continuing our conversation.

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

EMAIL #2: VALUE-ADD FOLLOW-UP (3-4 days after Email #1)

Subject: Resource for [Specific Challenge] at [Company Name]

Hi [First Name],

I hope this email finds you well. I was thinking about our recent conversation regarding [specific challenge], and I came across [article/case study/resource] that I thought would be valuable for you.

[Link to resource]

This [resource type] shows how [similar company] addressed [similar challenge] and achieved [specific results]. I thought the approach might be relevant to what you're trying to accomplish at [Company Name].

Have you had a chance to review the [materials/proposal] I sent in my last email? I'd love to hear your thoughts and answer any questions you might have.

Are you available for a brief call [this week/next week]?

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

EMAIL #3: CHECK-IN FOLLOW-UP (Week 2)

Subject: Checking in - [Company Name] and [Your Company]

Hi [First Name],

I wanted to check in and see if you've had a chance to review the [proposal/information/quote] I sent over. I know things can get busy, so I wanted to make sure this didn't get lost in your inbox.

QUICK RECAP:
We discussed how [Your Company] can help [Company Name]:
• [Benefit 1]
• [Benefit 2]
• [Benefit 3]

I'm here to answer any questions you might have. Would it be helpful to schedule a quick 15-minute call to address any concerns?

What does your calendar look like this week?

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

EMAIL #4: ADDRESSING OBJECTIONS (Week 3)

Subject: Addressing your questions about [Product/Service]

Hi [First Name],

I haven't heard back from you, and I'm wondering if there are any concerns or questions holding you back from moving forward.

COMMON QUESTIONS WE HEAR:
• "How long does implementation take?" - [Answer]
• "What kind of support do you provide?" - [Answer]
• "How does pricing compare?" - [Answer]

Is there something specific I can clarify? I'm happy to address any concerns you might have.

Alternatively, if now isn't the right time, I completely understand. Would it make sense to revisit this conversation in [timeframe]?

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

EMAIL #5: CASE STUDY FOLLOW-UP (Week 4)

Subject: How [Similar Company] achieved [Specific Result]

Hi [First Name],

I wanted to share a success story that's directly relevant to [Company Name].

[Similar Company], which faces similar challenges in [industry/area], recently implemented our [product/service]. They achieved:
• [Result 1 with specific metrics]
• [Result 2 with specific metrics]
• [Result 3 with specific metrics]

[Link to full case study]

I think you'd find their approach particularly interesting because [specific reason relevant to prospect].

Would you like to speak with them directly? I can arrange an introduction.

Let me know if you'd like to discuss how we could achieve similar results for [Company Name].

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

EMAIL #6: SPECIAL OFFER/URGENCY (Week 5-6)

Subject: Limited-time opportunity for [Company Name]

Hi [First Name],

I wanted to reach out one more time because we're offering [special promotion/limited-time offer] that could significantly benefit [Company Name].

SPECIAL OFFER:
• [Benefit 1 of offer]
• [Benefit 2 of offer]
• [Benefit 3 of offer]

This offer is available until [date], and I wanted to make sure you had the opportunity to take advantage of it.

Based on our earlier conversations, I believe this could help you [achieve specific goal] while [additional benefit].

Are you interested in learning more? I'm available for a call [day/time options].

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

EMAIL #7: BREAKUP EMAIL (Week 7-8)

Subject: Should I close your file?

Hi [First Name],

I've reached out several times regarding [product/service/solution] for [Company Name], but I haven't heard back from you.

I understand that:
• Timing might not be right
• Budget priorities may have shifted
• This might not be a priority anymore
• You're working with another solution

I don't want to be a pest, so I'm planning to close your file unless I hear from you.

If you're still interested, just reply to this email and we can continue our conversation. If not, I wish you and [Company Name] all the best.

Either way, I'd appreciate knowing where things stand.

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

POST-PROPOSAL FOLLOW-UP

Subject: Following up on our proposal for [Company Name]

Hi [First Name],

I wanted to follow up on the proposal I sent on [date] for [product/service].

PROPOSAL HIGHLIGHTS:
• Investment: $[Amount]
• Timeline: [Duration]
• Key Deliverables: [List]
• Expected ROI: [Percentage/Amount]

Have you had a chance to review it with your team? I'm here to answer any questions or address any concerns.

What would be the best next step from your perspective?

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

AFTER DEMO/PRESENTATION FOLLOW-UP

Subject: Thank you for your time - [Company Name] demo recap

Hi [First Name],

Thank you for attending the demo of [product/service] today. I hope you found it valuable and informative.

KEY FEATURES DISCUSSED:
• [Feature 1] - [Benefit]
• [Feature 2] - [Benefit]
• [Feature 3] - [Benefit]

ANSWERED QUESTIONS:
• [Question 1]: [Answer]
• [Question 2]: [Answer]

NEXT STEPS:
1. [Action item 1]
2. [Action item 2]
3. Schedule implementation call for [date]

I'm attaching [demo recording/presentation slides/additional resources] for your reference.

What questions can I answer for you?

Best regards,
[Your Name]
[Title]
[Company]
[Phone]
[Email]

---

TIPS FOR EFFECTIVE FOLLOW-UP:
• Personalize each email based on previous conversations
• Always provide value (resources, insights, case studies)
• Keep emails concise and scannable
• Include clear call-to-action
• Respect prospect's time and boundaries
• Track open rates and responses
• Adjust timing based on prospect's communication preferences
• Know when to stop following up
    `
  },
  {
    id: 'sales-customer-onboarding',
    title: 'Customer Onboarding Checklist',
    description: 'Comprehensive checklist to ensure smooth onboarding of new customers',
    category: 'sales',
    downloadCount: 3920,
    fileSize: '46.8 KB',
    rating: 4.9,
    tags: ['onboarding', 'customer', 'implementation', 'checklist', 'success'],
    createdAt: '2024-01-15T09:45:00Z',
    updatedAt: '2024-02-02T15:30:00Z',
    content: `
CUSTOMER ONBOARDING CHECKLIST

CUSTOMER INFORMATION:
Company Name: [Company Name]
Primary Contact: [Contact Name]
Email: [Email]
Phone: [Phone]
Account Manager: [Account Manager Name]
Start Date: [Date]
Contract Value: $[Amount]
Contract Term: [Duration]

PRE-ONBOARDING PREPARATION (Sales Team)
☐ Contract signed and executed
☐ Initial payment received
☐ Customer information collected
☐ Account created in system
☐ Internal handoff meeting scheduled
☐ Onboarding kickoff scheduled with customer
☐ Welcome email sent
☐ Resources and documentation prepared
☐ Implementation team assigned
☐ Success metrics defined
☐ Complete: _______ Date: _______

KICKOFF MEETING (Week 1)
☐ Welcome and introductions
☐ Review contract terms and deliverables
☐ Confirm customer goals and objectives
☐ Discuss timeline and milestones
☐ Assign roles and responsibilities
☐ Establish communication protocols
☐ Set expectations for response times
☐ Schedule regular check-in meetings
☐ Review technical requirements
☐ Provide access to customer portal
☐ Complete: _______ Date: _______

CUSTOMER GOALS & OBJECTIVES:
Primary Goal: [Goal]
Success Metrics:
• [Metric 1]
• [Metric 2]
• [Metric 3]
Timeline: [Timeline]

TECHNICAL SETUP (Week 1-2)
☐ Gather technical requirements
☐ System access provided
☐ User accounts created
☐ Permissions configured
☐ Integration requirements documented
☐ API keys generated (if applicable)
☐ Single Sign-On configured (if applicable)
☐ Data migration initiated (if needed)
☐ Security requirements verified
☐ Testing environment setup
☐ Complete: _______ Date: _______

TRAINING SCHEDULE (Week 2-3)
☐ Training needs assessment completed
☐ Training materials prepared
☐ Training sessions scheduled

Training Session 1: [Topic]
Date: [Date]
Duration: [Duration]
Attendees: [Names]
Status: ☐ Scheduled ☐ Completed

Training Session 2: [Topic]
Date: [Date]
Duration: [Duration]
Attendees: [Names]
Status: ☐ Scheduled ☐ Completed

Training Session 3: [Topic]
Date: [Date]
Duration: [Duration]
Attendees: [Names]
Status: ☐ Scheduled ☐ Completed

☐ Training recordings provided
☐ Quick reference guides shared
☐ FAQ document provided
☐ Complete: _______ Date: _______

DATA MIGRATION (if applicable)
☐ Data export from old system
☐ Data format verified
☐ Data mapping completed
☐ Test migration performed
☐ Data validation completed
☐ Full migration executed
☐ Data verification completed
☐ Old system access maintained (transition period)
☐ Complete: _______ Date: _______

CONFIGURATION & CUSTOMIZATION (Week 2-4)
☐ Workflow customization completed
☐ Report templates configured
☐ Notification settings configured
☐ Branding applied (if applicable)
☐ Custom fields created
☐ Integration testing completed
☐ User acceptance testing scheduled
☐ Customer feedback incorporated
☐ Final configuration approved
☐ Complete: _______ Date: _______

DOCUMENTATION PROVIDED:
☐ User manual
☐ Administrator guide
☐ API documentation
☐ Video tutorials
☐ FAQs
☐ Best practices guide
☐ Troubleshooting guide
☐ Support contact information
☐ Complete: _______ Date: _______

QUALITY ASSURANCE & TESTING (Week 3-4)
☐ Functionality testing completed
☐ Integration testing completed
☐ Performance testing completed
☐ Security testing completed
☐ User acceptance testing completed
☐ Issues documented and resolved
☐ Final approval from customer
☐ Complete: _______ Date: _______

GO-LIVE PREPARATION (Week 4)
☐ Go-live date confirmed
☐ Communication plan finalized
☐ Support team briefed
☐ Escalation process established
☐ Backup plan prepared
☐ Stakeholders notified
☐ Go-live checklist completed
☐ Complete: _______ Date: _______

GO-LIVE (Launch Day)
☐ System activated
☐ Users notified
☐ Support team on standby
☐ Monitor system performance
☐ Address immediate issues
☐ Collect user feedback
☐ Document lessons learned
☐ Complete: _______ Date: _______

POST GO-LIVE SUPPORT (Week 5-8)
☐ Daily check-ins (Week 1)
☐ Address outstanding issues
☐ Additional training provided (if needed)
☐ User adoption tracked
☐ Performance metrics reviewed
☐ Optimization recommendations provided
☐ Weekly check-ins (Week 2-4)
☐ Monthly check-ins scheduled
☐ Complete: _______ Date: _______

30-DAY REVIEW MEETING
Date: [Date]
Attendees: [Names]

Topics to Cover:
☐ Review success metrics
☐ Assess user adoption
☐ Identify challenges
☐ Discuss optimization opportunities
☐ Review support tickets
☐ Plan for continued success
☐ Gather feedback
☐ Complete: _______ Date: _______

60-DAY REVIEW MEETING
Date: [Date]
Attendees: [Names]

Topics to Cover:
☐ Review progress toward goals
☐ Analyze usage data
☐ Discuss advanced features
☐ Identify expansion opportunities
☐ Review satisfaction
☐ Plan next quarter
☐ Complete: _______ Date: _______

90-DAY REVIEW MEETING
Date: [Date]
Attendees: [Names]

Topics to Cover:
☐ Measure ROI
☐ Review achievement of success metrics
☐ Discuss long-term strategy
☐ Explore additional services
☐ Obtain testimonial/case study
☐ Schedule quarterly business reviews
☐ Complete: _______ Date: _______

CUSTOMER SUCCESS METRICS:

Usage Metrics:
• Active Users: [Number]
• Login Frequency: [Frequency]
• Feature Adoption: [Percentage]%
• Data Volume: [Amount]

Business Metrics:
• Goal Achievement: [Status]
• ROI: [Percentage/Amount]
• Efficiency Gains: [Metrics]
• Cost Savings: [Amount]

Satisfaction Metrics:
• Customer Satisfaction Score: [Score]
• Net Promoter Score: [Score]
• Support Ticket Volume: [Number]
• Response Time: [Time]

ONGOING ACCOUNT MANAGEMENT:
☐ Quarterly business reviews scheduled
☐ Account health monitoring
☐ Expansion opportunities identified
☐ Renewal process initiated (90 days before end of term)
☐ Continuous improvement plan
☐ Relationship building activities

HANDOFF TO CUSTOMER SUCCESS:
☐ Account documentation complete
☐ Customer profile updated
☐ Success plan documented
☐ Key contacts identified
☐ Relationship history recorded
☐ Upsell/cross-sell opportunities noted
☐ Formal handoff meeting completed
☐ Complete: _______ Date: _______

CUSTOMER FEEDBACK:
[Space for recording customer feedback throughout onboarding]

LESSONS LEARNED:
[Space for recording insights for future onboarding improvements]

SIGNATURES:

Account Manager: _________________ Date: _______
[Name]

Implementation Manager: _________________ Date: _______
[Name]

Customer Contact: _________________ Date: _______
[Name, Title]
    `
  },
  {
    id: 'sales-sla-agreement',
    title: 'Service Level Agreement (SLA)',
    description: 'Professional SLA template defining service standards and commitments',
    category: 'sales',
    downloadCount: 2890,
    fileSize: '54.2 KB',
    rating: 4.7,
    tags: ['sla', 'service-level', 'agreement', 'support', 'commitments'],
    createdAt: '2024-01-18T10:30:00Z',
    updatedAt: '2024-02-03Т12:45:00Z',
    content: `
SERVICE LEVEL AGREEMENT (SLA)

This Service Level Agreement ("SLA") is entered into on [Date] between:

SERVICE PROVIDER:
[Company Name] ("Provider")
[Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]

CLIENT:
[Company Name] ("Client")
[Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]

1. AGREEMENT OVERVIEW
This SLA defines the service levels, performance standards, and responsibilities between the Provider and Client for the provision of [services description].

1.1 Effective Date: [Date]
1.2 Term: [Duration]
1.3 Review Period: [Quarterly/Annually]

2. SERVICES PROVIDED
The Provider agrees to deliver the following services:

2.1 Core Services:
• [Service 1]: [Description]
• [Service 2]: [Description]
• [Service 3]: [Description]

2.2 Service Hours:
• Business Hours: [Hours] [Time Zone]
• Business Days: [Days]
• After-Hours Support: [Available/Not Available]

2.3 Service Exclusions:
• [Exclusion 1]
• [Exclusion 2]
• [Exclusion 3]

3. SERVICE LEVEL OBJECTIVES

3.1 SYSTEM AVAILABILITY
• Target Uptime: [Percentage]% measured monthly
• Planned Maintenance: [Hours] per month with [Notice Period] advance notice
• Maximum Planned Downtime: [Hours] per month
• Measurement Period: 24/7 basis
• Exclusions: Scheduled maintenance, client-caused outages, force majeure

Uptime Calculation:
Uptime % = ((Total Minutes in Month - Downtime Minutes) / Total Minutes in Month) × 100

3.2 RESPONSE TIMES

Support Ticket Priority Levels:

PRIORITY 1 - CRITICAL:
• Definition: Complete service outage or critical business impact
• Initial Response Time: [Minutes/Hours]
• Target Resolution Time: [Hours]
• Communication Frequency: Every [Hours]
• Escalation: Immediate to senior management

PRIORITY 2 - HIGH:
• Definition: Major functionality impaired, significant business impact
• Initial Response Time: [Hours]
• Target Resolution Time: [Hours/Days]
• Communication Frequency: Every [Hours/Days]
• Escalation: After [Hours]

PRIORITY 3 - MEDIUM:
• Definition: Minor functionality impaired, moderate business impact
• Initial Response Time: [Hours/Days]
• Target Resolution Time: [Days]
• Communication Frequency: Daily updates
• Escalation: After [Days]

PRIORITY 4 - LOW:
• Definition: Minimal impact, general inquiries
• Initial Response Time: [Days]
• Target Resolution Time: [Days]
• Communication Frequency: As needed
• Escalation: After [Days]

3.3 PERFORMANCE STANDARDS
• System Response Time: [Seconds] for [Percentage]% of transactions
• Page Load Time: [Seconds] average
• API Response Time: [Milliseconds] average
• Data Backup Frequency: [Daily/Weekly]
• Backup Retention: [Duration]
• Disaster Recovery Time: [Hours/Days]

3.4 QUALITY STANDARDS
• First Contact Resolution Rate: [Percentage]%
• Customer Satisfaction Score: [Score] out of 10
• Ticket Resolution Rate: [Percentage]% within SLA
• Defect Rate: Less than [Percentage]%

4. SUPPORT SERVICES

4.1 Support Channels:
• Phone: [Phone Number]
• Email: [Email Address]
• Online Portal: [URL]
• Chat: [Available/Not Available]

4.2 Support Team:
• Dedicated Account Manager: [Yes/No]
• Technical Support Team: [24/7/Business Hours]
• Escalation Contact: [Name/Contact]

4.3 Support Scope:
• Installation and configuration assistance
• Troubleshooting and problem resolution
• Software updates and patches
• User training and documentation
• Performance optimization
• [Additional support services]

5. MONITORING AND REPORTING

5.1 System Monitoring:
• 24/7 automated monitoring
• Proactive issue detection
• Real-time alerts
• Performance dashboards

5.2 Regular Reports:
• Monthly Performance Report: Due by [Day] of following month
• Quarterly Business Review: Scheduled [Frequency]
• Annual Service Review: Scheduled [Timeframe]

5.3 Report Contents:
• Availability statistics
• Response and resolution times
• Incident summary
• Performance metrics
• Improvement recommendations

6. CLIENT RESPONSIBILITIES

The Client agrees to:
• Provide accurate contact information
• Designate primary and backup contacts
• Provide timely access to systems and information
• Follow documented procedures and best practices
• Report issues promptly
• Maintain compatible infrastructure
• Comply with acceptable use policies
• Provide feedback on service quality

7. MAINTENANCE AND UPDATES

7.1 Scheduled Maintenance:
• Frequency: [Weekly/Monthly]
• Maintenance Window: [Day/Time]
• Advance Notice: [Hours/Days]
• Maximum Duration: [Hours]

7.2 Emergency Maintenance:
• Performed when critical issues threaten service
• Advance notice provided when possible
• Client notified as soon as practical

7.3 Updates and Upgrades:
• Software Updates: [Frequency]
• Security Patches: Within [Hours/Days] of release
• Major Upgrades: [Frequency] with [Notice Period] notice
• Client testing period: [Days] before production deployment

8. ESCALATION PROCEDURES

Level 1: Support Team
Contact: [Email/Phone]
Responsibility: Initial response and resolution

Level 2: Technical Lead
Contact: [Email/Phone]
Escalation Trigger: [Condition]
Response Time: [Hours]

Level 3: Management
Contact: [Email/Phone]
Escalation Trigger: [Condition]
Response Time: [Hours]

Level 4: Executive
Contact: [Email/Phone]
Escalation Trigger: [Condition]
Response Time: [Hours]

9. SERVICE CREDITS

If Provider fails to meet SLA commitments, Client may be eligible for service credits:

9.1 Availability Credits:
• [Percentage range]% uptime: [Credit percentage]% of monthly fee
• [Percentage range]% uptime: [Credit percentage]% of monthly fee
• Below [Percentage]% uptime: [Credit percentage]% of monthly fee

9.2 Response Time Credits:
• [Percentage]% of tickets outside response SLA: [Credit]% of monthly fee

9.3 Credit Request Process:
• Client must request credits within [Days] of incident
• Request must include ticket numbers and dates
• Provider will review and respond within [Days]
• Credits applied to next month's invoice
• Maximum total credits: [Percentage]% of monthly fee

9.4 Credit Exclusions:
Credits not applicable for outages caused by:
• Client's actions or equipment
• Force majeure events
• Scheduled maintenance
• Third-party service failures
• Internet connectivity issues
• Factors outside Provider's control

10. CHANGE MANAGEMENT

10.1 Service Changes:
• Either party may request service changes
• Changes require mutual written agreement
• Impact assessment provided within [Days]
• Implementation timeline agreed upon
• Documentation updated accordingly

10.2 SLA Modifications:
• SLA reviewed [Frequency]
• Modifications require [Notice Period] notice
• Changes documented in writing
• Both parties must approve modifications

11. SECURITY AND COMPLIANCE

11.1 Security Standards:
• [Security certifications/standards]
• Data encryption in transit and at rest
• Regular security audits
• Vulnerability assessments
• Penetration testing [Frequency]

11.2 Compliance:
• [Industry regulations - e.g., GDPR, HIPAA, SOC 2]
• Compliance audits [Frequency]
• Documentation provided upon request

11.3 Data Protection:
• Data backup: [Frequency]
• Backup retention: [Duration]
• Disaster recovery: [RTO/RPO]
• Data breach notification: Within [Hours]

12. LIMITATION OF LIABILITY

12.1 Provider's liability limited to:
• Direct damages only
• Maximum of [Amount] or [Percentage]% of annual contract value
• Exclusions: indirect, consequential, or punitive damages

12.2 Exceptions:
• Intentional misconduct
• Gross negligence
• Breach of confidentiality
• [Other exceptions]

13. TERMINATION

13.1 Termination for Cause:
Either party may terminate for material breach with [Notice Period] written notice if breach not cured.

13.2 Termination for Convenience:
Client may terminate with [Notice Period] notice and payment of [Terms].

13.3 Upon Termination:
• Provider will assist with transition for [Days]
• Client data returned within [Days]
• Final invoice issued
• Service credits applied

14. GENERAL PROVISIONS

14.1 Entire Agreement: This SLA constitutes the entire agreement regarding service levels.

14.2 Governing Law: [State/Country]

14.3 Dispute Resolution: [Mediation/Arbitration process]

14.4 Force Majeure: Neither party liable for delays due to circumstances beyond reasonable control.

14.5 Notices: All notices in writing to addresses specified above.

15. ACCEPTANCE

By signing below, both parties agree to the terms of this Service Level Agreement.

SERVICE PROVIDER:               CLIENT:

_________________________      _________________________
[Name], [Title]                [Name], [Title]
[Company Name]                 [Company Name]
Date: _______________          Date: _______________

APPENDICES:

Appendix A: Detailed Service Descriptions
Appendix B: Performance Metrics and Measurement Methods
Appendix C: Contact List and Escalation Matrix
Appendix D: Acceptable Use Policy
Appendix E: Change Request Form
    `
  },
  {
    id: 'sales-renewal-proposal',
    title: 'Renewal Proposal Template',
    description: 'Professional proposal template for contract renewals with updated terms',
    category: 'sales',
    downloadCount: 3340,
    fileSize: '47.5 KB',
    rating: 4.8,
    tags: ['renewal', 'proposal', 'contract', 'retention', 'upsell'],
    createdAt: '2024-01-20T13:15:00Z',
    updatedAt: '2024-02-04T11:30:00Z',
    content: `
CONTRACT RENEWAL PROPOSAL

[Company Name]
[Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]
[Website]

PREPARED FOR:
[Client Company Name]
[Client Name, Title]
[Client Address]
[City, State, ZIP Code]

PROPOSAL DATE: [Date]
PREPARED BY: [Account Manager Name]
VALID UNTIL: [Date]

EXECUTIVE SUMMARY

Dear [Client Name],

Thank you for being a valued client of [Company Name]. As your current contract approaches its expiration date of [Date], we are pleased to present this renewal proposal for your continued partnership.

Over the past [contract period], we have been honored to serve [Client Company] and support your [specific achievements/goals]. We look forward to continuing our successful partnership and helping you achieve even greater success in the coming year.

CURRENT CONTRACT OVERVIEW

Current Agreement Details:
• Contract Start Date: [Date]
• Contract End Date: [Date]
• Current Services: [List services]
• Current Annual Investment: $[Amount]
• Contract Term: [Duration]

PARTNERSHIP HIGHLIGHTS & ACHIEVEMENTS

During our partnership, we have achieved the following together:

Key Metrics:
• [Metric 1]: [Result]
• [Metric 2]: [Result]
• [Metric 3]: [Result]
• [Metric 4]: [Result]

Notable Accomplishments:
• [Achievement 1]
• [Achievement 2]
• [Achievement 3]

Value Delivered:
• Total ROI: [Percentage/Amount]
• Cost Savings: $[Amount]
• Efficiency Gains: [Metric]
• [Other value metrics]

RENEWAL PROPOSAL OPTIONS

We are pleased to offer the following renewal options:

---

OPTION 1: STANDARD RENEWAL
Continuation of Current Services

Services Included:
• [Service 1]
• [Service 2]
• [Service 3]
• [Service 4]

Contract Term: [Duration]
Annual Investment: $[Amount]
Monthly Payment: $[Amount]

Benefits:
• Maintain current service levels
• [Benefit 1]
• [Benefit 2]
• Price lock for contract term

---

OPTION 2: ENHANCED PACKAGE (RECOMMENDED)
Expanded Services for Growing Needs

Services Included:
All Standard Services PLUS:
• [Additional Service 1]
• [Additional Service 2]
• [Additional Service 3]
• [Upgraded feature 1]
• [Upgraded feature 2]

Contract Term: [Duration]
Annual Investment: $[Amount]
Monthly Payment: $[Amount]

Benefits:
• [Enhanced benefit 1]
• [Enhanced benefit 2]
• [Enhanced benefit 3]
• Priority support
• Dedicated account manager
• Quarterly business reviews

Added Value: $[Amount] (XX% savings vs. buying separately)

---

OPTION 3: PREMIUM PACKAGE
Comprehensive Solution for Maximum Impact

Services Included:
All Enhanced Services PLUS:
• [Premium Service 1]
• [Premium Service 2]
• [Premium Service 3]
• [Premium Feature 1]
• [Premium Feature 2]
• Custom integrations
• Advanced analytics

Contract Term: [Duration]
Annual Investment: $[Amount]
Monthly Payment: $[Amount]

Benefits:
• [Premium benefit 1]
• [Premium benefit 2]
• [Premium benefit 3]
• 24/7 Priority Support
• Executive Sponsor
• Monthly strategy sessions
• Custom reporting
• Advanced training

Added Value: $[Amount] (XX% savings vs. buying separately)

---

RENEWAL INCENTIVES

As a valued existing client, we are offering special renewal incentives:

Early Renewal Bonus:
Renew by [Date] and receive:
• [Incentive 1]
• [Incentive 2]
• [Discount percentage]% discount
• [Additional months] free

Multi-Year Commitment:
• 2-Year Term: [Discount]% discount
• 3-Year Term: [Discount]% discount

Referral Program:
• Refer new clients and receive [benefit]

WHAT'S NEW & IMPROVED

Since your last contract, we have:

Product Enhancements:
• [Enhancement 1]
• [Enhancement 2]
• [Enhancement 3]

New Features:
• [Feature 1]: [Description]
• [Feature 2]: [Description]
• [Feature 3]: [Description]

Service Improvements:
• [Improvement 1]
• [Improvement 2]
• [Improvement 3]

PRICING COMPARISON

Service Component | Current Contract | Proposed Renewal | Change
------------------|------------------|------------------|--------
[Component 1] | $[Amount] | $[Amount] | [%]
[Component 2] | $[Amount] | $[Amount] | [%]
[Component 3] | $[Amount] | $[Amount] | [%]
TOTAL | $[Amount] | $[Amount] | [%]

Value Justification:
[Explanation of any pricing changes]

IMPLEMENTATION & TRANSITION

Renewal Timeline:
• [Date]: Proposal presented
• [Date]: Decision deadline for early renewal bonus
• [Date]: Current contract expires
• [Date]: Renewal contract begins

Transition Plan:
• Seamless continuation of services
• No disruption to current operations
• [Any system updates/migrations]
• Account review meeting scheduled

YOUR DEDICATED TEAM

Account Manager: [Name]
Email: [Email]
Phone: [Phone]

Technical Lead: [Name]
Email: [Email]
Phone: [Phone]

Customer Success Manager: [Name]
Email: [Email]
Phone: [Phone]

YEAR AHEAD STRATEGIC PLAN

Goals for Next Contract Period:
• [Goal 1]
• [Goal 2]
• [Goal 3]

Recommended Initiatives:
• [Initiative 1]: [Description and expected impact]
• [Initiative 2]: [Description and expected impact]
• [Initiative 3]: [Description and expected impact]

Success Metrics:
• [Metric 1 and target]
• [Metric 2 and target]
• [Metric 3 and target]

TERMS & CONDITIONS

Payment Terms:
• Payment Schedule: [Monthly/Quarterly/Annual]
• Payment Methods: [Methods accepted]
• Late Payment: [Terms]

Contract Terms:
• Auto-renewal: [Yes/No with terms]
• Cancellation: [Notice period and terms]
• Price Protection: [Terms]
• Service Level Agreement: [Reference to SLA]

Included:
• [Included item 1]
• [Included item 2]
• [Included item 3]

Not Included:
• [Excluded item 1]
• [Excluded item 2]
• [Excluded item 3]

FREQUENTLY ASKED QUESTIONS

Q: Why are prices changing?
A: [Answer]

Q: Can I modify my services mid-contract?
A: [Answer]

Q: What happens if I don't renew?
A: [Answer]

Q: Can I switch between packages?
A: [Answer]

Q: What if my needs change during the contract?
A: [Answer]

TESTIMONIALS

"[Client testimonial about your service]"
- [Name, Title, Company]

"[Another client testimonial]"
- [Name, Title, Company]

NEXT STEPS

To accept this renewal proposal:

1. Review the options and select your preferred package
2. Sign the renewal agreement (attached)
3. Return signed agreement via:
   • Email: [Email]
   • DocuSign: [Link]
   • Mail: [Address]
4. Submit payment or payment authorization

Questions? Let's discuss:
• Schedule a call: [Calendly link]
• Email: [Email]
• Phone: [Phone]

We value your partnership and look forward to continuing to serve [Client Company] in achieving your goals.

Thank you for your continued trust in [Company Name].

Sincerely,

[Your Name]
[Title]
[Company Name]
[Phone]
[Email]

ACCEPTANCE

I accept the renewal proposal as follows:

Selected Package:
☐ Option 1: Standard Renewal
☐ Option 2: Enhanced Package
☐ Option 3: Premium Package

Contract Term:
☐ 1 Year
☐ 2 Years ([Discount]% discount)
☐ 3 Years ([Discount]% discount)

Client Signature: _________________ Date: _______
Print Name: [Name]
Title: [Title]
Company: [Company Name]

ATTACHMENTS:
• Detailed Service Description
• Service Level Agreement
• Terms and Conditions
• Current vs. Proposed Feature Comparison
• Case Studies
• Product Roadmap
    `
  },
  {
    id: 'sales-commission-sheet',
    title: 'Sales Commission Sheet',
    description: 'Comprehensive tracking sheet for calculating sales commissions and bonuses',
    category: 'sales',
    downloadCount: 4670,
    fileSize: '49.1 KB',
    rating: 4.6,
    tags: ['commission', 'compensation', 'tracking', 'sales-rep', 'incentive'],
    createdAt: '2024-01-22T11:00:00Z',
    updatedAt: '2024-02-05T14:45:00Z',
    content: `
SALES COMMISSION TRACKING SHEET

COMPANY: [Company Name]
COMMISSION PERIOD: [Month/Quarter/Year]
SALES REPRESENTATIVE: [Sales Rep Name]
EMPLOYEE ID: [ID Number]
TERRITORY: [Territory Name]
MANAGER: [Manager Name]

COMMISSION STRUCTURE SUMMARY:

Base Salary: $[Amount] per [Month/Year]
Commission Rate: [Percentage]% of sales
Quota: $[Amount] per [Period]
Bonus Opportunity: $[Amount] at 100% quota attainment

Commission Tiers:
• 0-74% of Quota: [Rate]% commission
• 75-99% of Quota: [Rate]% commission
• 100-124% of Quota: [Rate]% commission
• 125%+ of Quota: [Rate]% commission

SALES ACTIVITY TRACKING:

Deal # | Date Closed | Customer Name | Product/Service | Sale Amount | Commission Rate | Commission Amount | Payment Status | Notes
-------|-------------|---------------|-----------------|-------------|-----------------|-------------------|----------------|-------
[#] | [Date] | [Customer] | [Product] | $[Amount] | [%] | $[Amount] | [Status] | [Notes]
[#] | [Date] | [Customer] | [Product] | $[Amount] | [%] | $[Amount] | [Status] | [Notes]
[#] | [Date] | [Customer] | [Product] | $[Amount] | [%] | $[Amount] | [Status] | [Notes]
[#] | [Date] | [Customer] | [Product] | $[Amount] | [%] | $[Amount] | [Status] | [Notes]
[#] | [Date] | [Customer] | [Product] | $[Amount] | [%] | $[Amount] | [Status] | [Notes]

MONTHLY PERFORMANCE SUMMARY:

Total Sales: $[Amount]
Quota: $[Amount]
Quota Attainment: [Percentage]%
Base Commission Earned: $[Amount]
Bonus/Accelerators: $[Amount]
Adjustments: $[Amount]
Total Commission: $[Amount]

DETAILED COMMISSION CALCULATIONS:

Product Category Performance:
Category | Sales Amount | Commission Rate | Commission Earned
---------|--------------|-----------------|------------------
[Category 1] | $[Amount] | [%] | $[Amount]
[Category 2] | $[Amount] | [%] | $[Amount]
[Category 3] | $[Amount] | [%] | $[Amount]
TOTAL | $[Amount] | | $[Amount]

New Business vs. Renewals:
Type | Sales Amount | Commission Rate | Commission Earned
-----|--------------|-----------------|------------------
New Business | $[Amount] | [%] | $[Amount]
Renewals | $[Amount] | [%] | $[Amount]
Upsells | $[Amount] | [%] | $[Amount]
TOTAL | $[Amount] | | $[Amount]

BONUS CALCULATIONS:

Quota Attainment Bonus:
Quota: $[Amount]
Actual Sales: $[Amount]
Attainment: [Percentage]%
Bonus Rate: [Percentage]% of base salary
Bonus Amount: $[Amount]

Performance Accelerators:
☐ Exceeded quota by 25%+: $[Bonus Amount]
☐ Largest deal of quarter: $[Bonus Amount]
☐ New market penetration: $[Bonus Amount]
☐ Customer satisfaction >90%: $[Bonus Amount]
☐ [Custom accelerator]: $[Bonus Amount]

Total Bonus: $[Amount]

SPIFF & CONTEST EARNINGS:
Contest/SPIFF Name | Criteria | Amount Earned | Status
-------------------|----------|---------------|--------
[Contest 1] | [Criteria] | $[Amount] | [Won/Pending]
[Contest 2] | [Criteria] | $[Amount] | [Won/Pending]
[Contest 3] | [Criteria] | $[Amount] | [Won/Pending]
TOTAL SPIFFS | | $[Amount] |

DEDUCTIONS & ADJUSTMENTS:

Reason | Amount | Approval | Notes
-------|--------|----------|-------
Cancelled Deal: [Customer] | $([Amount]) | [Manager] | [Notes]
Customer Return: [Customer] | $([Amount]) | [Manager] | [Notes]
Commission Recovery | $([Amount]) | [Manager] | [Notes]
Administrative Adjustment | $([Amount]) | [Manager] | [Notes]
TOTAL DEDUCTIONS | $([Amount]) | |

COMMISSION RECONCILIATION:

Gross Commission Earned: $[Amount]
Bonuses & Accelerators: $[Amount]
SPIFFs & Contests: $[Amount]
Subtotal: $[Amount]
Deductions: $([Amount])
Adjustments: $[Amount]

TOTAL COMMISSION PAYMENT: $[Amount]

PAYMENT DETAILS:

Payment Date: [Date]
Payment Method: ☐ Direct Deposit ☐ Check
Payment Period: [Period]
Check Number: [Number]
Tax Withholding: $[Amount]
Net Payment: $[Amount]

YEAR-TO-DATE SUMMARY:

Period | Sales | Quota | Attainment | Commission | Bonus | Total Earnings
-------|-------|-------|------------|------------|-------|---------------
Q1 | $[Amount] | $[Amount] | [%] | $[Amount] | $[Amount] | $[Amount]
Q2 | $[Amount] | $[Amount] | [%] | $[Amount] | $[Amount] | $[Amount]
Q3 | $[Amount] | $[Amount] | [%] | $[Amount] | $[Amount] | $[Amount]
Q4 | $[Amount] | $[Amount] | [%] | $[Amount] | $[Amount] | $[Amount]
YTD TOTAL | $[Amount] | $[Amount] | [%] | $[Amount] | $[Amount] | $[Amount]

PIPELINE VALUE:
Current Quarter Pipeline: $[Amount]
Next Quarter Pipeline: $[Amount]
Weighted Pipeline: $[Amount]
Projected Commission: $[Amount]

KEY PERFORMANCE INDICATORS:

Metric | Target | Actual | Status
-------|--------|--------|--------
Sales Volume | $[Amount] | $[Amount] | [Status]
Number of Deals | [Number] | [Number] | [Status]
Average Deal Size | $[Amount] | $[Amount] | [Status]
Win Rate | [%] | [%] | [Status]
Sales Cycle (days) | [Days] | [Days] | [Status]
Customer Retention | [%] | [%] | [Status]

COMMISSION DISPUTES OR QUESTIONS:

Date | Issue Description | Resolution | Resolved By | Date Resolved
-----|-------------------|------------|-------------|---------------
[Date] | [Issue] | [Resolution] | [Name] | [Date]
[Date] | [Issue] | [Resolution] | [Name] | [Date]

NOTES & COMMENTS:
[Space for additional notes, special circumstances, or explanations]

APPROVAL & SIGNATURES:

Sales Representative Acknowledgment:
I acknowledge that the commission calculation above is accurate and agree to the payment amount.

Sales Rep Signature: _________________ Date: _______
[Sales Rep Name]

Sales Manager Approval:
I have reviewed and approve the commission calculation for this period.

Manager Signature: _________________ Date: _______
[Manager Name]

Finance/Accounting Approval:
Commission verified and approved for payment.

Finance Rep Signature: _________________ Date: _______
[Finance Rep Name]

Date Processed: [Date]

COMMISSION PLAN TERMS & CONDITIONS:

• Commissions paid [Frequency] following month-end close
• Sales must be closed and invoiced to earn commission
• Customer payment required for commission payout [Yes/No]
• Cancelled deals subject to commission recovery
• Returns/refunds may result in commission chargebacks
• Split deals: commission divided [per agreement]
• All commissions subject to management approval
• Commission plan subject to change with [Notice Period] notice
• Disputes must be raised within [Days] of statement

For questions regarding this statement, contact:
[Finance Contact Name]
[Email]
[Phone]
    `
  },
  {
    id: 'sales-price-list',
    title: 'Price List & Rate Card',
    description: 'Professional price list template with product catalog and pricing tiers',
    category: 'sales',
    downloadCount: 5890,
    fileSize: '43.7 KB',
    rating: 4.7,
    tags: ['pricing', 'rate-card', 'price-list', 'catalog', 'products'],
    createdAt: '2024-01-25T14:45:00Z',
    updatedAt: '2024-02-06T09:30:00Z',
    content: `
PRICE LIST & RATE CARD

[Company Name]
[Address]
[City, State, ZIP Code]
[Phone Number]
[Email Address]
[Website]

EFFECTIVE DATE: [Date]
VALID THROUGH: [Date]
VERSION: [Version Number]

GENERAL TERMS:

• All prices in [Currency]
• Prices subject to change without notice
• Minimum order: $[Amount] or [Quantity] units
• Volume discounts available
• Custom quotes available for large orders
• Payment terms: [Net 30/Net 15/Due on Receipt]
• Shipping not included unless specified
• Tax not included

PRODUCT CATALOG

---

CATEGORY 1: [PRODUCT CATEGORY NAME]

Product Code | Product Name | Description | Unit | List Price | Quantity Discount
-------------|--------------|-------------|------|------------|------------------
[SKU-001] | [Product Name] | [Brief description] | [Each] | $[Price] | 10+: [%] off, 50+: [%] off
[SKU-002] | [Product Name] | [Brief description] | [Each] | $[Price] | 10+: [%] off, 50+: [%] off
[SKU-003] | [Product Name] | [Brief description] | [Each] | $[Price] | 10+: [%] off, 50+: [%] off
[SKU-004] | [Product Name] | [Brief description] | [Each] | $[Price] | 10+: [%] off, 50+: [%] off

---

CATEGORY 2: [PRODUCT CATEGORY NAME]

Product Code | Product Name | Description | Unit | List Price | Quantity Discount
-------------|--------------|-------------|------|------------|------------------
[SKU-101] | [Product Name] | [Brief description] | [Each] | $[Price] | 10+: [%] off, 50+: [%] off
[SKU-102] | [Product Name] | [Brief description] | [Each] | $[Price] | 10+: [%] off, 50+: [%] off
[SKU-103] | [Product Name] | [Brief description] | [Each] | $[Price] | 10+: [%] off, 50+: [%] off
[SKU-104] | [Product Name] | [Brief description] | [Each] | $[Price] | 10+: [%] off, 50+: [%] off

---

SERVICE OFFERINGS

Service Name | Description | Rate/Price | Unit | Terms
-------------|-------------|------------|------|-------
[Service 1] | [Detailed description] | $[Rate] | [Per hour/project/month] | [Minimum/terms]
[Service 2] | [Detailed description] | $[Rate] | [Per hour/project/month] | [Minimum/terms]
[Service 3] | [Detailed description] | $[Rate] | [Per hour/project/month] | [Minimum/terms]
[Service 4] | [Detailed description] | $[Rate] | [Per hour/project/month] | [Minimum/terms]

---

PROFESSIONAL SERVICES RATE CARD

Role/Level | Hourly Rate | Daily Rate | Monthly Rate | Expertise
-----------|-------------|------------|--------------|----------
Junior Consultant | $[Rate] | $[Rate] | $[Rate] | [1-2 years experience]
Consultant | $[Rate] | $[Rate] | $[Rate] | [3-5 years experience]
Senior Consultant | $[Rate] | $[Rate] | $[Rate] | [6-10 years experience]
Principal Consultant | $[Rate] | $[Rate] | $[Rate] | [10+ years experience]
Subject Matter Expert | $[Rate] | $[Rate] | $[Rate] | [Specialized expertise]
Project Manager | $[Rate] | $[Rate] | $[Rate] | [Certified PM]

Minimum Engagement: [Hours/Days]
After-Hours Rate: [Rate] × regular rate
Weekend Rate: [Rate] × regular rate
Holiday Rate: [Rate] × regular rate
Travel Time: Billed at [Percentage]% of regular rate

---

SUBSCRIPTION PACKAGES

STARTER PACKAGE
Monthly: $[Amount]
Annual: $[Amount] (Save [%]!)

Features:
• [Feature 1]
• [Feature 2]
• [Feature 3]
• [Feature 4]
• [Feature 5]

Limits:
• [Limit 1]
• [Limit 2]
• Support: Email only, 48-hour response

Best For: [Target customer description]

---

PROFESSIONAL PACKAGE (MOST POPULAR)
Monthly: $[Amount]
Annual: $[Amount] (Save [%]!)

Features:
Everything in Starter, PLUS:
• [Additional Feature 1]
• [Additional Feature 2]
• [Additional Feature 3]
• [Additional Feature 4]
• [Additional Feature 5]

Limits:
• [Limit 1]
• [Limit 2]
• Support: Email & phone, 24-hour response

Best For: [Target customer description]

---

ENTERPRISE PACKAGE
Monthly: $[Amount]
Annual: $[Amount] (Save [%]!)

Features:
Everything in Professional, PLUS:
• [Premium Feature 1]
• [Premium Feature 2]
• [Premium Feature 3]
• Unlimited [Feature]
• Dedicated account manager
• Custom integrations
• Priority support
• SLA guarantee

Limits:
• Unlimited users
• Unlimited [resource]
• Support: 24/7 phone & email, 4-hour response

Best For: [Target customer description]

---

CUSTOM ENTERPRISE
Contact Sales for Quote

Fully customized solution including:
• Custom development
• White-label options
• Dedicated infrastructure
• Onsite training
• 24/7 dedicated support
• Custom SLA
• Volume pricing

Contact: [Sales Email/Phone]

---

ADD-ON SERVICES & FEATURES

Add-On | Description | Price | Billing
-------|-------------|-------|--------
[Add-on 1] | [Description] | $[Amount] | [Per month/one-time]
[Add-on 2] | [Description] | $[Amount] | [Per month/one-time]
[Add-on 3] | [Description] | $[Amount] | [Per month/one-time]
[Add-on 4] | [Description] | $[Amount] | [Per month/one-time]
[Add-on 5] | [Description] | $[Amount] | [Per month/one-time]

---

VOLUME DISCOUNT SCHEDULE

Annual Contract Value | Discount
----------------------|----------
$[Amount] - $[Amount] | [Percentage]%
$[Amount] - $[Amount] | [Percentage]%
$[Amount] - $[Amount] | [Percentage]%
$[Amount]+ | [Percentage]% + custom terms

Quantity Discounts:
• 10-24 units: [Percentage]% off
• 25-49 units: [Percentage]% off
• 50-99 units: [Percentage]% off
• 100+ units: [Percentage]% off + free shipping

---

IMPLEMENTATION & TRAINING SERVICES

Service | Description | Price | Duration
--------|-------------|-------|----------
Basic Setup | [Description] | $[Amount] | [Hours/Days]
Standard Implementation | [Description] | $[Amount] | [Hours/Days]
Premium Implementation | [Description] | $[Amount] | [Hours/Days]
Enterprise Implementation | [Description] | Custom Quote | [Timeframe]
On-site Training (per day) | [Description] | $[Amount] | 8 hours
Virtual Training (per session) | [Description] | $[Amount] | 2-4 hours
Custom Development (per hour) | [Description] | $[Amount] | As needed
Data Migration | [Description] | $[Amount] | [Hours/Days]

---

SUPPORT & MAINTENANCE PLANS

Basic Support (Included with all packages)
• Email support
• Knowledge base access
• Community forum
• [Response time]

Standard Support - $[Amount]/month
• Everything in Basic
• Phone support
• [Response time]
• [Hours of coverage]
• Monthly system health check

Premium Support - $[Amount]/month
• Everything in Standard
• Priority response
• [Response time]
• 24/7 coverage
• Dedicated support engineer
• Quarterly business reviews
• Custom reporting

---

SHIPPING & HANDLING

Shipping Method | Domestic | International | Delivery Time
----------------|----------|---------------|---------------
Standard Ground | $[Amount] | $[Amount] | [Days]
2-Day Express | $[Amount] | $[Amount] | [Days]
Next Day | $[Amount] | $[Amount] | [Days]
Freight (pallets) | Quote | Quote | [Days]

Free Shipping: Orders over $[Amount] (domestic only)

---

SPECIAL OFFERS & PROMOTIONS

Current Promotions:
• [Promotion 1]: [Description and discount]
  Valid: [Start Date] - [End Date]
  Code: [Promo Code]

• [Promotion 2]: [Description and discount]
  Valid: [Start Date] - [End Date]
  Code: [Promo Code]

• [Promotion 3]: [Description and discount]
  Valid: [Start Date] - [End Date]
  Code: [Promo Code]

---

PAYMENT OPTIONS

Accepted Payment Methods:
• Credit Card (Visa, Mastercard, Amex, Discover)
• ACH/Bank Transfer
• Wire Transfer
• Purchase Order (approved accounts)
• PayPal
• Check (with approved credit)

Payment Terms:
• New customers: Payment in advance or COD
• Approved customers: Net 30
• Subscription services: Auto-pay monthly/annually
• Late payment fee: [Percentage]% per month

---

REFUND & RETURN POLICY

• 30-day money-back guarantee on [products/services]
• Return shipping: Customer responsibility
• Restocking fee: [Percentage]% on [items]
• Opened/used items: No returns
• Custom orders: No returns
• Services: Refundable based on work completed
• Subscription cancellation: [Terms]

---

WARRANTY INFORMATION

Standard Warranty:
• Duration: [Period]
• Coverage: [What's covered]
• Exclusions: [What's not covered]

Extended Warranty Available:
• [Duration]: $[Amount]
• Covers: [Additional coverage]

---

TERMS & CONDITIONS

• Prices subject to change without notice
• We reserve the right to correct pricing errors
• All sales final unless otherwise noted
• Products subject to availability
• Lead times are estimates, not guarantees
• Custom quotes valid for [Days]
• Bulk orders may require advance notice
• International orders subject to customs duties
• All prices exclude taxes unless noted
• See full Terms & Conditions at [URL]

---

CONTACT INFORMATION

Sales Inquiries:
Phone: [Phone Number]
Email: [Sales Email]
Hours: [Business Hours] [Time Zone]

Customer Support:
Phone: [Phone Number]
Email: [Support Email]
Hours: [Support Hours] [Time Zone]

Website: [URL]
Online Ordering: [URL]
Live Chat: Available [Hours]

Request a Custom Quote:
For custom pricing, bulk orders, or enterprise solutions:
• Fill out form at: [URL]
• Email: [Enterprise Sales Email]
• Call: [Direct Sales Number]

Response time: Within [Hours/Days]

---

FOOTNOTES

1. Volume discounts calculated on single orders only
2. Annual subscriptions paid in full upfront
3. Enterprise pricing requires minimum [commitment]
4. Setup fees may apply for new accounts
5. Overage charges apply for usage beyond plan limits
6. Price increases limited to [%] annually for existing contracts
7. Educational and non-profit discounts available upon verification

Last Updated: [Date]
Price List Version: [Version]

For the most current pricing, visit: [Website URL]
    `
  }
];

// Add sales templates to the main templates array
templates.push(...salesTemplates);

// Marketing Templates
const marketingTemplates: DocumentTemplate[] = [
  {
    id: 'marketing-social-media-calendar',
    title: 'Social Media Content Calendar',
    description: 'Comprehensive calendar for planning and organizing social media content across multiple platforms',
    category: 'marketing',
    downloadCount: 6780,
    fileSize: '48.9 KB',
    rating: 4.9,
    tags: ['social-media', 'content-calendar', 'planning', 'scheduling', 'marketing'],
    createdAt: '2024-01-07T09:00:00Z',
    updatedAt: '2024-01-28T15:30:00Z',
    content: `
SOCIAL MEDIA CONTENT CALENDAR

CAMPAIGN OVERVIEW:
Campaign Name: [Campaign Name]
Month/Quarter: [Time Period]
Marketing Manager: [Name]
Last Updated: [Date]

SOCIAL MEDIA ACCOUNTS:
• Facebook: [Account Name/URL]
• Instagram: [Account Name/URL]
• Twitter/X: [Account Name/URL]
• LinkedIn: [Account Name/URL]
• TikTok: [Account Name/URL]
• YouTube: [Account Name/URL]
• Pinterest: [Account Name/URL]

CAMPAIGN GOALS:
• Primary Goal: [Goal]
• Target Audience: [Audience description]
• Key Messages: [Key messages]
• Success Metrics: [KPIs]

CONTENT THEMES FOR THE MONTH:
Week 1: [Theme]
Week 2: [Theme]
Week 3: [Theme]
Week 4: [Theme]

WEEKLY POSTING SCHEDULE:

WEEK 1 - [Date Range]

Monday [Date]:
Platform | Time | Content Type | Post Copy | Visual/Media | Link | Hashtags | Status
---------|------|--------------|-----------|--------------|------|----------|--------
Facebook | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Instagram | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Twitter | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
LinkedIn | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]

Tuesday [Date]:
Platform | Time | Content Type | Post Copy | Visual/Media | Link | Hashtags | Status
---------|------|--------------|-----------|--------------|------|----------|--------
Facebook | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Instagram | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Twitter | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
LinkedIn | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]

Wednesday [Date]:
Platform | Time | Content Type | Post Copy | Visual/Media | Link | Hashtags | Status
---------|------|--------------|-----------|--------------|------|----------|--------
Facebook | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Instagram | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Twitter | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
LinkedIn | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]

Thursday [Date]:
Platform | Time | Content Type | Post Copy | Visual/Media | Link | Hashtags | Status
---------|------|--------------|-----------|--------------|------|----------|--------
Facebook | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Instagram | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Twitter | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
LinkedIn | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]

Friday [Date]:
Platform | Time | Content Type | Post Copy | Visual/Media | Link | Hashtags | Status
---------|------|--------------|-----------|--------------|------|----------|--------
Facebook | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Instagram | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Twitter | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
LinkedIn | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]

Saturday [Date]:
Platform | Time | Content Type | Post Copy | Visual/Media | Link | Hashtags | Status
---------|------|--------------|-----------|--------------|------|----------|--------
Facebook | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Instagram | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]

Sunday [Date]:
Platform | Time | Content Type | Post Copy | Visual/Media | Link | Hashtags | Status
---------|------|--------------|-----------|--------------|------|----------|--------
Instagram | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]
Facebook | [Time] | [Type] | [Copy preview] | [Image/Video desc] | [URL] | [#tags] | [Draft/Scheduled/Posted]

CONTENT TYPES:
• Promotional Post
• Educational Content
• Behind-the-Scenes
• User-Generated Content
• Testimonial/Review
• Product Showcase
• Blog Post Share
• Video Content
• Infographic
• Poll/Question
• Live Event Announcement
• Story/Reel
• Carousel Post
• Trending Topic

OPTIMAL POSTING TIMES BY PLATFORM:
Facebook: [e.g., 1-4 PM Wed-Fri]
Instagram: [e.g., 11 AM-1 PM Mon-Thu]
Twitter: [e.g., 8 AM-10 AM, 6 PM-9 PM daily]
LinkedIn: [e.g., 7-8 AM, 12 PM, 5-6 PM Tue-Thu]
TikTok: [e.g., 6-10 AM, 7-11 PM]

HASHTAG STRATEGY:

Brand Hashtags (Always include):
• [#YourBrand]
• [#YourProductLine]
• [#BrandTagline]

Campaign Hashtags:
• [#CampaignName]
• [#SeasonalPromo]

Trending/Industry Hashtags (Rotate):
• [#IndustryTerm]
• [#TrendingTopic]
• [#CommunityHashtag]

Platform-Specific Guidelines:
• Instagram: 10-15 hashtags
• Twitter: 1-2 hashtags
• LinkedIn: 3-5 hashtags
• TikTok: 4-8 hashtags

CONTENT PILLARS:

Pillar 1: Education (40% of content)
Topics: [Industry tips, how-tos, tutorials]
Goal: [Build authority and provide value]

Pillar 2: Inspiration (25% of content)
Topics: [Success stories, motivation, behind-the-scenes]
Goal: [Connect emotionally with audience]

Pillar 3: Promotion (20% of content)
Topics: [Products, services, offers]
Goal: [Drive conversions and sales]

Pillar 4: Community (15% of content)
Topics: [User-generated content, engagement posts, polls]
Goal: [Build relationships and loyalty]

CONTENT LIBRARY:

Asset Type | File Name | Platform | Usage Rights | Date Created | Notes
-----------|-----------|----------|--------------|--------------|-------
Image | [filename.jpg] | [All/Specific] | [Owned/Licensed] | [Date] | [Notes]
Video | [filename.mp4] | [All/Specific] | [Owned/Licensed] | [Date] | [Notes]
Graphic | [filename.png] | [All/Specific] | [Owned/Licensed] | [Date] | [Notes]
Template | [filename.psd] | [All/Specific] | [Owned/Licensed] | [Date] | [Notes]

SPECIAL EVENTS & HOLIDAYS:

Date | Event/Holiday | Planned Content | Platforms | Lead Time | Status
-----|---------------|-----------------|-----------|-----------|--------
[Date] | [Event name] | [Content description] | [Platforms] | [Days ahead] | [Status]
[Date] | [Event name] | [Content description] | [Platforms] | [Days ahead] | [Status]
[Date] | [Event name] | [Content description] | [Platforms] | [Days ahead] | [Status]

ENGAGEMENT STRATEGY:

Response Time Goals:
• Comments: [Within X hours]
• Direct Messages: [Within X hours]
• Mentions: [Within X hours]

Community Management:
• Monitor brand mentions daily
• Engage with user-generated content
• Respond to all comments (positive and negative)
• Share/repost relevant content from followers

Influencer Partnerships:
• [Influencer 1]: [Collaboration details]
• [Influencer 2]: [Collaboration details]

PAID PROMOTION SCHEDULE:

Date Range | Platform | Campaign | Budget | Target Audience | Objective | Status
-----------|----------|----------|--------|-----------------|-----------|--------
[Dates] | [Platform] | [Campaign name] | $[Amount] | [Demographics/interests] | [Reach/Engagement/Conversions] | [Status]
[Dates] | [Platform] | [Campaign name] | $[Amount] | [Demographics/interests] | [Reach/Engagement/Conversions] | [Status]

PERFORMANCE METRICS TO TRACK:

Engagement Metrics:
• Likes: [Target per post]
• Comments: [Target per post]
• Shares: [Target per post]
• Saves: [Target per post]
• Engagement Rate: [Target %]

Growth Metrics:
• Follower Growth: [Target per month]
• Reach: [Target impressions]
• Profile Visits: [Target]

Conversion Metrics:
• Click-Through Rate: [Target %]
• Website Clicks: [Target]
• Conversions: [Target]
• Revenue Attributed: $[Target]

WEEKLY ANALYTICS REVIEW:

Week Ending [Date]:
Platform | Posts | Reach | Engagement | Top Post | CTR | New Followers
---------|-------|-------|------------|----------|-----|---------------
Facebook | [#] | [#] | [#] | [Description] | [%] | [#]
Instagram | [#] | [#] | [#] | [Description] | [%] | [#]
Twitter | [#] | [#] | [#] | [Description] | [%] | [#]
LinkedIn | [#] | [#] | [#] | [Description] | [%] | [#]

Key Insights:
• [Insight 1]
• [Insight 2]
• [Insight 3]

Actions for Next Week:
• [Action 1]
• [Action 2]
• [Action 3]

COMPETITOR ANALYSIS:

Competitor | Platforms | Posting Freq | Top Content | Engagement Rate | Strategy Notes
-----------|-----------|--------------|-------------|-----------------|----------------
[Competitor 1] | [Platforms] | [X/week] | [Type] | [%] | [Observations]
[Competitor 2] | [Platforms] | [X/week] | [Type] | [%] | [Observations]
[Competitor 3] | [Platforms] | [X/week] | [Type] | [%] | [Observations]

CONTENT APPROVAL WORKFLOW:

1. Content Creation: [Team member name]
2. First Review: [Team member name]
3. Final Approval: [Team member name]
4. Scheduling: [Team member name]
5. Post-Publication Monitoring: [Team member name]

Approval Deadlines:
• Content due: [X days before posting]
• First review: [X days before posting]
• Final approval: [X days before posting]
• Scheduled: [X days before posting]

NOTES & IDEAS:

Content Ideas to Develop:
• [Idea 1]
• [Idea 2]
• [Idea 3]

Lessons Learned:
• [Learning 1]
• [Learning 2]
• [Learning 3]

Upcoming Opportunities:
• [Opportunity 1]
• [Opportunity 2]

Prepared by: _________________ Date: _______
[Marketing Manager Name]
    `
  },
  {
    id: 'marketing-blog-post-outline',
    title: 'Blog Post Outline Template',
    description: 'Structured template for planning and outlining engaging blog posts with SEO optimization',
    category: 'marketing',
    downloadCount: 5420,
    fileSize: '42.3 KB',
    rating: 4.8,
    tags: ['blog', 'content', 'writing', 'seo', 'outline'],
    createdAt: '2024-01-09T10:30:00Z',
    updatedAt: '2024-01-30T12:45:00Z',
    content: `
BLOG POST OUTLINE TEMPLATE

POST INFORMATION:
Working Title: [Blog Post Title]
Author: [Author Name]
Target Publish Date: [Date]
Status: [Draft/In Review/Approved/Scheduled/Published]
Blog Category: [Category]
Tags: [Tag1, Tag2, Tag3, Tag4]
Word Count Goal: [1500/2000/2500/3000] words

CONTENT STRATEGY:

Primary Goal:
☐ Drive organic traffic
☐ Generate leads
☐ Build thought leadership
☐ Educate audience
☐ Promote product/service
☐ Improve brand awareness
☐ Other: [Specify]

Target Audience:
Primary Persona: [Persona name]
• Demographics: [Age, role, industry]
• Knowledge Level: [Beginner/Intermediate/Advanced]
• Pain Points: [Specific challenges this post addresses]
• Reading Intent: [What they hope to gain]

SEO STRATEGY:

Primary Keyword: [main keyword phrase]
Search Volume: [monthly searches]
Keyword Difficulty: [score/10]
Current Ranking: [position or "not ranking"]

Secondary Keywords:
• [Secondary keyword 1]
• [Secondary keyword 2]
• [Secondary keyword 3]

Long-Tail Keywords:
• [Long-tail variation 1]
• [Long-tail variation 2]
• [Long-tail variation 3]

Search Intent:
☐ Informational (seeking knowledge)
☐ Navigational (looking for specific site/page)
☐ Commercial (researching before purchase)
☐ Transactional (ready to take action)

Competitor Analysis:
Top 3 Ranking Articles:
1. [URL] - [Strengths/Gaps]
2. [URL] - [Strengths/Gaps]
3. [URL] - [Strengths/Gaps]

Our Differentiation:
[How this post will be better/different]

META DATA & SEO ELEMENTS:

SEO Title (50-60 characters):
[Compelling title with primary keyword]

Meta Description (150-160 characters):
[Engaging description with primary keyword and call-to-action]

URL Slug:
/blog/[keyword-rich-slug]

Featured Image:
Description: [What the image shows]
Alt Text: [Descriptive alt text with keyword]
Dimensions: [1200x630 for social sharing]
Source: [Stock photo site/original/licensed]

HEADLINE OPTIONS:

Option 1: [Headline variation 1]
Option 2: [Headline variation 2]
Option 3: [Headline variation 3]
Option 4: [Headline variation 4]

Selected Headline: [Final headline choice]

Subheadline (Optional):
[Supporting headline that adds context]

INTRODUCTION (150-200 words)

Hook (1-2 sentences):
[Attention-grabbing opening - statistic, question, bold statement, or story]

Problem/Pain Point:
[Clearly state the problem or challenge the reader faces]

Promise/Solution Preview:
[What the reader will learn or gain by reading this post]

Credibility Statement:
[Why you/your company is qualified to address this topic]

Thesis Statement:
[Main argument or central point of the blog post]

Preview of Main Points:
In this post, you'll discover:
• [Key point 1]
• [Key point 2]
• [Key point 3]

MAIN CONTENT OUTLINE:

H2: [Section 1 Heading - Include keyword variation]
Opening paragraph: [Brief introduction to this section]

H3: [Subpoint 1.1]
• Key point: [Main idea]
• Supporting details: [Explanation]
• Example: [Specific example or scenario]
• Data/Statistic: [Relevant data with source]

H3: [Subpoint 1.2]
• Key point: [Main idea]
• Supporting details: [Explanation]
• Quote: "[Expert quote]" - [Source, Title]
• Visual: [Describe chart/image needed]

H3: [Subpoint 1.3]
• Key point: [Main idea]
• Supporting details: [Explanation]
• Case study: [Brief description of example]
• Takeaway: [Key lesson]

---

H2: [Section 2 Heading - Include keyword variation]
Opening paragraph: [Brief introduction to this section]

H3: [Subpoint 2.1]
• Key point: [Main idea]
• Supporting details: [Explanation]
• List:
  1. [Item 1]
  2. [Item 2]
  3. [Item 3]

H3: [Subpoint 2.2]
• Key point: [Main idea]
• How-to steps:
  Step 1: [Action]
  Step 2: [Action]
  Step 3: [Action]
• Pro tip: [Insider advice]

H3: [Subpoint 2.3]
• Key point: [Main idea]
• Common mistake: [What to avoid]
• Best practice: [Recommended approach]
• Tool/Resource: [Helpful tool mention]

---

H2: [Section 3 Heading - Include keyword variation]
Opening paragraph: [Brief introduction to this section]

H3: [Subpoint 3.1]
• Key point: [Main idea]
• Supporting details: [Explanation]
• Real-world example: [Specific scenario]
• Visual: [Infographic/diagram description]

H3: [Subpoint 3.2]
• Key point: [Main idea]
• Supporting details: [Explanation]
• Comparison: [X vs. Y analysis]
• Recommendation: [Specific advice]

H3: [Subpoint 3.3]
• Key point: [Main idea]
• Supporting details: [Explanation]
• Checklist:
  ☐ [Item 1]
  ☐ [Item 2]
  ☐ [Item 3]

---

[Add additional H2 sections as needed]

SUPPORTING ELEMENTS:

Statistics & Data Points:
• [Statistic 1] - Source: [URL]
• [Statistic 2] - Source: [URL]
• [Statistic 3] - Source: [URL]

Expert Quotes:
• "[Quote 1]" - [Expert Name, Title, Company]
• "[Quote 2]" - [Expert Name, Title, Company]

Examples & Case Studies:
• [Company/Person]: [Brief description of their success/failure]
• [Company/Person]: [Brief description of their success/failure]

Visual Content Needed:
☐ Featured image
☐ Header image for introduction
☐ [Infographic about X topic]
☐ [Screenshot showing Y]
☐ [Diagram illustrating Z process]
☐ [Chart/graph with data]
☐ [Comparison table]
☐ [Pull quote graphic]

INTERNAL LINKS (3-5 recommended):

• [Anchor text 1] → [URL to related blog post]
• [Anchor text 2] → [URL to pillar content]
• [Anchor text 3] → [URL to product/service page]
• [Anchor text 4] → [URL to resource page]
• [Anchor text 5] → [URL to about/contact page]

EXTERNAL LINKS (2-4 recommended):

• [Anchor text 1] → [URL to authoritative source]
• [Anchor text 2] → [URL to research/study]
• [Anchor text 3] → [URL to tool/resource]
• [Anchor text 4] → [URL to industry publication]

CONCLUSION (150-200 words)

Summary of Main Points:
Recap the key takeaways:
• [Main point 1]
• [Main point 2]
• [Main point 3]

Final Thought/Insight:
[Memorable closing statement or future outlook]

Call-to-Action (CTA):

Primary CTA:
[What you want the reader to do next]

Options:
☐ Download [resource name]
☐ Sign up for [newsletter/webinar/trial]
☐ Request [consultation/demo]
☐ Read [related article]
☐ Share on social media
☐ Leave a comment
☐ Try [product/service]

CTA Copy:
[Compelling call-to-action text]

CTA Button/Link Text:
[Action-oriented button text]

CTA Link URL:
[Destination URL]

ADDITIONAL COMPONENTS:

Key Takeaways Box (Optional):
Place near top or bottom of post
• [Takeaway 1]
• [Takeaway 2]
• [Takeaway 3]
• [Takeaway 4]

FAQ Section (Optional but recommended for SEO):
Q: [Common question 1]
A: [Clear, concise answer]

Q: [Common question 2]
A: [Clear, concise answer]

Q: [Common question 3]
A: [Clear, concise answer]

Table of Contents (For posts 2000+ words):
• [Link to section 1]
• [Link to section 2]
• [Link to section 3]

Author Bio:
[Author name] is [brief credentials and expertise]. [One sentence about their background]. [Link to author page or social profile]

Related Posts (3-5):
• [Related post title 1]
• [Related post title 2]
• [Related post title 3]

Lead Magnet/Content Upgrade (Optional):
[Downloadable checklist/template/guide related to post topic]

PROMOTION STRATEGY:

Email Newsletter:
Subject Line: [Subject]
Preview Text: [Preview]
Email Excerpt: [First 2-3 paragraphs]
CTA: [Read More button]
Send Date: [Date]
Segment: [Audience segment]

Social Media Posts:

LinkedIn:
Post Copy: [Copy with hook]
Visual: [Image/graphic]
Hashtags: [#tag1 #tag2 #tag3]
Post Time: [Day/Time]

Twitter/X:
Tweet 1: [Copy with link]
Tweet 2: [Pull quote + link]
Tweet 3: [Statistic + link]
Hashtags: [#tag1 #tag2]
Schedule: [Day/Time for each]

Facebook:
Post Copy: [Copy]
Visual: [Image]
Post Time: [Day/Time]

Instagram:
Caption: [Copy]
Visual: [Image design]
Hashtags: [Up to 30 hashtags]
Story: [Yes/No + design]
Post Time: [Day/Time]

Paid Promotion:
☐ Google Ads (budget: $[amount])
☐ Facebook/Instagram Ads (budget: $[amount])
☐ LinkedIn Sponsored Content (budget: $[amount])
☐ Reddit Ads (budget: $[amount])
Target Audience: [Demographics/interests]

Outreach & Distribution:
• Email to [influencer name] for potential share
• Submit to [content aggregator]
• Share in [LinkedIn/Facebook group]
• Post in [Slack community]
• Send to [industry newsletter]

RESEARCH SOURCES:

Primary Sources:
1. [Source title/URL] - [Key information to use]
2. [Source title/URL] - [Key information to use]
3. [Source title/URL] - [Key information to use]

Secondary Sources:
• [Source title/URL]
• [Source title/URL]
• [Source title/URL]

Tools/Data Sources:
• [Tool/database used for research]
• [Survey/report referenced]

CONTENT CHECKLIST:

Pre-Writing:
☐ Keyword research completed
☐ Competitor content analyzed
☐ Outline reviewed and approved
☐ Sources and data gathered
☐ Visuals planned

Writing:
☐ Introduction written (includes hook, problem, promise)
☐ All H2/H3 sections completed
☐ Conclusion written with strong CTA
☐ Word count target achieved
☐ Primary keyword used naturally (avoid keyword stuffing)
☐ Secondary keywords incorporated
☐ Examples and data included
☐ Internal links added (3-5)
☐ External links added (2-4)
☐ Transition sentences between sections

Editing & Optimization:
☐ Proofread for grammar and spelling
☐ Checked for readability (Hemingway/Grammarly)
☐ Sentences and paragraphs kept short
☐ Removed jargon or explained technical terms
☐ Active voice used (passive minimized)
☐ Formatted with bullets, lists, bold text
☐ Mobile-friendly formatting verified

SEO Elements:
☐ Primary keyword in title (preferably at beginning)
☐ Primary keyword in first 100 words
☐ Primary keyword in URL slug
☐ Primary keyword in meta description
☐ Keywords in at least 2-3 H2 headings
☐ Images added with descriptive alt text
☐ Featured image optimized (compressed)
☐ Meta title is 50-60 characters
☐ Meta description is 150-160 characters
☐ Internal link structure implemented
☐ External authoritative sources linked
☐ Content length competitive with top-ranking posts
☐ Schema markup added (Article, FAQ, HowTo if applicable)

Visual Elements:
☐ Featured image created/selected
☐ All images optimized and compressed
☐ Alt text added to all images
☐ Infographics/charts created
☐ Screenshots captured and annotated
☐ Images properly attributed/licensed

Final Review:
☐ Fact-checking completed
☐ All links tested (no broken links)
☐ CTA included and link works
☐ Author bio added
☐ Related posts suggested
☐ Categories and tags assigned
☐ Preview on mobile and desktop
☐ Legal/compliance review (if needed)
☐ Final approval received

Post-Publication:
☐ Post published at scheduled time
☐ Shared on all social media channels
☐ Email newsletter sent
☐ Submitted to search console for indexing
☐ Shared in relevant communities
☐ Outreach emails sent
☐ Analytics tracking verified
☐ Monitor for comments and engagement
☐ Respond to comments within 24 hours
☐ Track performance metrics

PERFORMANCE TRACKING:

Success Metrics:
• Organic traffic: [Target monthly visitors]
• Time on page: [Target minutes:seconds]
• Bounce rate: [Target %]
• Social shares: [Target number]
• Comments: [Target number]
• Backlinks: [Target number]
• Email signups: [Target number]
• Downloads: [Target number]
• Keyword rankings: [Target positions]
• Conversion rate: [Target %]

Tracking Tools:
• Google Analytics
• Google Search Console
• Social media analytics
• Email marketing platform
• Rank tracking tool

Review Schedule:
• 1 week post-publication: Initial performance review
• 1 month: Full performance analysis
• 3 months: Optimization review
• 6 months: Update/refresh assessment

Optimization Opportunities:
☐ Update with new data/statistics
☐ Add new sections based on questions/comments
☐ Improve underperforming CTAs
☐ Add more visuals
☐ Update internal links
☐ Refresh introduction
☐ Target new keywords

NOTES & IDEAS:

Content Spin-Off Ideas:
• [Related topic 1 for future post]
• [Related topic 2 for future post]
• [Infographic version]
• [Video script adaptation]
• [Social media carousel]

Questions to Address:
• [Question raised during research]
• [Question from audience]

Additional Resources Needed:
• [Resource 1]
• [Resource 2]

APPROVALS:

Created by: _________________ Date: _______
[Author Name]

Reviewed by: _________________ Date: _______
[Editor Name]

SEO Approved by: _________________ Date: _______
[SEO Specialist Name]

Final Approval by: _________________ Date: _______
[Marketing Manager Name]
    `
  },
  {
    id: 'marketing-email-campaign',
    title: 'Email Campaign Template',
    description: 'Complete email marketing campaign template with strategy, copy framework, and tracking',
    category: 'marketing',
    downloadCount: 7340,
    fileSize: '46.7 KB',
    rating: 4.9,
    tags: ['email', 'campaign', 'marketing', 'newsletter', 'automation'],
    createdAt: '2024-01-11T13:15:00Z',
    updatedAt: '2024-02-01T09:30:00Z',
    content: `
EMAIL CAMPAIGN TEMPLATE

CAMPAIGN OVERVIEW:
Campaign Name: [Campaign Name]
Campaign Type: ☐ Newsletter ☐ Promotional ☐ Drip Sequence ☐ Welcome Series ☐ Re-engagement ☐ Event ☐ Product Launch
Campaign Owner: [Name]
Date Created: [Date]
Launch Date: [Date]
End Date: [Date if applicable]

CAMPAIGN OBJECTIVES:

Primary Goal:
☐ Generate sales/revenue
☐ Drive website traffic
☐ Nurture leads
☐ Increase engagement
☐ Build brand awareness
☐ Event registrations
☐ Content downloads
☐ Collect feedback/reviews
☐ Other: [Specify]

Success Metrics (SMART Goals):
• Open Rate Target: [%] (Industry avg: 15-25%)
• Click-Through Rate Target: [%] (Industry avg: 2-5%)
• Conversion Rate Target: [%]
• Revenue Target: $[Amount]
• List Growth: [Number of new subscribers]
• Unsubscribe Rate Limit: < [%] (Keep under 0.5%)

TARGET AUDIENCE:

Segment Name: [Segment identifier]
List Size: [Number of recipients]

Demographics:
• Age Range: [Range]
• Location: [Geographic targeting]
• Gender: [If relevant]
• Job Title/Role: [If B2B]
• Income Level: [If relevant]
• Company Size: [If B2B]

Behavioral Criteria:
• Purchase history: [Recent buyers, never purchased, etc.]
• Website activity: [Pages visited, time on site]
• Email engagement: [Openers, clickers, inactive]
• Lead score: [Range]
• Cart abandonment: [Yes/No]
• Product interests: [Categories/products]

Psychographics:
• Interests: [List relevant interests]
• Pain points: [Key challenges]
• Goals: [What they want to achieve]
• Values: [What matters to them]

CAMPAIGN STRATEGY:

Email Sequence Structure:
This campaign consists of [#] emails sent over [timeframe]

Email #1: [Purpose/Name]
Trigger/Send Date: [Date or trigger event]
Send Time: [Time + Timezone]
Delay from previous: [N/A or X days/hours]

Email #2: [Purpose/Name]
Trigger/Send Date: [Date or trigger event]
Send Time: [Time + Timezone]
Delay from previous: [X days/hours]

Email #3: [Purpose/Name]
Trigger/Send Date: [Date or trigger event]
Send Time: [Time + Timezone]
Delay from previous: [X days/hours]

[Add more emails as needed]

---

EMAIL #1: [EMAIL NAME/PURPOSE]

SUBJECT LINE:

Option 1: [Subject line variation 1]
Option 2: [Subject line variation 2]
Option 3: [Subject line variation 3]

Selected Subject Line: [Final choice]

Subject Line Best Practices Check:
☐ Under 50 characters (mobile-friendly)
☐ Includes personalization token ({{First_Name}})
☐ Creates curiosity or urgency
☐ Clear value proposition
☐ No spam trigger words (FREE, !!!, ACT NOW)
☐ Action-oriented
☐ A/B test variant created

PREVIEW TEXT (Preheader):
[40-100 characters that complement subject line]

SENDER INFORMATION:
From Name: [Company Name] or [Personal Name from Company]
From Email: [email@company.com]
Reply-To Email: [reply@company.com]

EMAIL HEADER:
• Logo: [Company logo with link to homepage]
• Header Image: [Optional hero image]
• Navigation: [Home | Shop | Blog | Contact] (optional)

EMAIL BODY CONTENT:

Greeting:
Hi {{First_Name}},

or

Hello,

or

Hi there,

Opening Hook (First paragraph):
[Attention-grabbing opening that relates to recipient's interests or pain points. Make it personal and relevant. 2-3 sentences max.]

Value Proposition:
[Clear statement of what's in it for them. Why should they keep reading?]

Main Content:

[SECTION 1 HEADING]
[Content paragraph explaining the main message. Keep paragraphs short - 2-3 sentences max.]

Key Points (use bullets for scannability):
• [Benefit/Feature 1]
• [Benefit/Feature 2]
• [Benefit/Feature 3]

[SECTION 2 HEADING - if needed]
[Additional content paragraph]

Social Proof (if applicable):
"[Customer testimonial or impressive statistic]"
- [Customer Name, Company]

or

[X,XXX] customers have already [achieved result]

Visual Elements:
• Hero Image: [Description]
• Product Image: [Description]
• Icon/Graphic: [Description]
• GIF/Animation: [Description]

Primary Call-to-Action (CTA):

CTA Button Text: [Action verb + benefit, e.g., "Get My Free Guide"]
Button Color: [Color - high contrast]
Button Size: [Large, thumb-friendly]
Link URL: [Full URL with UTM parameters]
Placement: [After which section]

CTA Supporting Copy:
[1-2 sentences before CTA button that reinforce value and create urgency]

Secondary CTA (Optional):
Link Text: [e.g., "Learn more about X"]
Link URL: [URL]
Placement: [Location in email]

Closing Paragraph:
[Friendly closing that reinforces the CTA or adds a personal touch]

Signature:
Best regards,
[Name]
[Title]
[Company Name]

or

Cheers,
The [Company Name] Team

P.S. [Powerful postscript that reinforces key benefit, adds urgency, or includes additional offer]

EMAIL FOOTER:

Contact Information:
[Company Name]
[Street Address]
[City, State ZIP Code]
[Phone Number]

Social Media Links:
[Facebook Icon] [Twitter Icon] [LinkedIn Icon] [Instagram Icon]

Footer Links:
• [Unsubscribe]
• [Update Email Preferences]
• [View in Browser]
• [Privacy Policy]
• [Contact Us]

Legal/Compliance Text:
"You're receiving this email because you [signed up/made a purchase/attended event]. If you no longer wish to receive these emails, you can [unsubscribe link] at any time."

"[Company Name], [Full Address]"

---

EMAIL #2: [EMAIL NAME/PURPOSE]

[Repeat structure for second email]

Subject Line: [Subject]
Preview Text: [Preview]
[Full email copy following same structure]

---

EMAIL #3: [EMAIL NAME/PURPOSE]

[Repeat structure for third email]

---

A/B TESTING STRATEGY:

Test Element:
☐ Subject line
☐ From name (Company vs. Person)
☐ Preview text
☐ Email copy (short vs. long)
☐ CTA button text
☐ CTA button color
☐ Send time
☐ Personalization level
☐ Image vs. text-only
☐ Number of CTAs (single vs. multiple)

Test Details:
Variant A: [Description]
Variant B: [Description]

Sample Size: [% of list, e.g., 10% to A, 10% to B]
Test Duration: [How long before declaring winner]
Winner Criteria: [Which metric determines winner]
Winner Rollout: Send to remaining [80%] of list

TECHNICAL SETUP:

Email Service Provider: [Platform name - Mailchimp, HubSpot, etc.]
Template ID: [Template identifier]
Campaign ID: [ID number]
List/Segment ID: [ID]

Personalization Tokens:
• {{First_Name}}
• {{Company_Name}}
• {{Last_Purchase_Product}}
• {{Days_Since_Purchase}}
• {{Custom_Field}}

Dynamic Content Blocks:
• [Show Block A to Segment X]
• [Show Block B to Segment Y]
• [Product recommendations based on browse history]
• [Location-specific content]

Tracking & UTM Parameters:
URL structure: [URL]?utm_source=[source]&utm_medium=email&utm_campaign=[campaign_name]&utm_content=[content_identifier]

Example:
https://yoursite.com/product?utm_source=mailchimp&utm_medium=email&utm_campaign=summer_sale_2024&utm_content=cta_button

Conversion Tracking:
☐ Goal URL set up
☐ Event tracking configured
☐ Revenue tracking enabled
☐ Email platform connected to analytics

MOBILE OPTIMIZATION:

Requirements:
☐ Responsive design (single column)
☐ Minimum 14px font size
☐ CTA buttons at least 44x44px (thumb-friendly)
☐ Images scale to fit screen
☐ Compressed images (<200KB total)
☐ Avoid small tap targets close together
☐ Test on iOS Mail app
☐ Test on Gmail app (Android)
☐ Test on Outlook mobile

DELIVERABILITY CHECKLIST:

Authentication:
☐ SPF record verified
☐ DKIM signature enabled
☐ DMARC policy set
☐ Domain reputation checked

List Hygiene:
☐ Hard bounces removed
☐ Unsubscribes processed
☐ Invalid emails removed
☐ Inactive subscribers (1+ year) segmented
☐ Re-engagement campaign sent to inactive users

Pre-Send Testing:
☐ Spam score checked (<5 recommended)
☐ All links tested and working
☐ Images loading properly
☐ Alt text added to all images
☐ Plain text version created
☐ Tested in Gmail
☐ Tested in Outlook
☐ Tested in Apple Mail
☐ Tested in Yahoo Mail
☐ Mobile preview checked (iOS and Android)
☐ Personalization tokens tested
☐ Unsubscribe link tested
☐ Legal compliance verified (CAN-SPAM/GDPR)

Content Quality:
☐ Text-to-image ratio balanced (60:40)
☐ No spam trigger words in subject/body
☐ Not using all caps or excessive punctuation!!!
☐ Link count reasonable (<10 links)
☐ Proper HTML structure
☐ No broken formatting

AUTOMATION TRIGGERS (If applicable):

Trigger Event:
☐ User signs up
☐ Purchase made
☐ Cart abandoned
☐ Downloads resource
☐ Clicks specific link
☐ Visits specific page
☐ Reaches lead score threshold
☐ Birthday/Anniversary
☐ Subscription renewal date
☐ X days of inactivity
☐ Tag added/removed

Trigger Conditions:
[Specific conditions that must be met]

Wait Time:
Send email [immediately / X hours / X days] after trigger event

Exit Conditions:
User exits sequence if:
• [Condition 1, e.g., Makes a purchase]
• [Condition 2, e.g., Unsubscribes]
• [Condition 3, e.g., Opens competitor email]

CAMPAIGN BUDGET:

Email Design/Template: $[Amount]
Copywriting: $[Amount]
Images/Graphics: $[Amount]
Email Service Provider: $[Monthly cost]
A/B Testing Tools: $[Amount]
List Cleaning Service: $[Amount]
Total Campaign Cost: $[Amount]

Expected ROI: [Revenue target] / [Total cost] = [X]x ROI

TIMELINE & MILESTONES:

[Date] - Campaign strategy approved
[Date] - Copy drafted
[Date] - Design completed
[Date] - Internal review
[Date] - Revisions completed
[Date] - Final approval
[Date] - Setup in ESP
[Date] - Testing completed
[Date] - Campaign launched
[Date] - Mid-campaign check
[Date] - Campaign concludes
[Date] - Performance review meeting

RISK MITIGATION:

Potential Risks & Backup Plans:
• Low open rate: [Have subject line variants ready]
• High unsubscribe rate: [Pause send, review content]
• Technical issues: [Have backup send date]
• Deliverability problems: [Monitor and adjust]
• Negative feedback: [Response protocol ready]

PERFORMANCE TRACKING:

Delivery Metrics:
• Emails Sent: [Target #]
• Delivered: [#]
• Bounced: [#]
• Bounce Rate: [%]

Engagement Metrics:
• Opens: [#]
• Open Rate: [%]
• Unique Opens: [#]
• Clicks: [#]
• Click-Through Rate (CTR): [%]
• Click-to-Open Rate (CTOR): [%]

Conversion Metrics:
• Conversions: [#]
• Conversion Rate: [%]
• Revenue Generated: $[Amount]
• Average Order Value: $[Amount]
• Revenue Per Email: $[Amount]

List Health:
• Unsubscribes: [#]
• Unsubscribe Rate: [%]
• Spam Complaints: [#]
• Complaint Rate: [%]

Engagement by Segment:
Segment | Opens | CTR | Conversions | Revenue
--------|-------|-----|-------------|--------
[Segment 1] | [%] | [%] | [#] | $[Amount]
[Segment 2] | [%] | [%] | [#] | $[Amount]
[Segment 3] | [%] | [%] | [#] | $[Amount]

Device Stats:
• Desktop: [%]
• Mobile: [%]
• Tablet: [%]

Email Client Stats:
• Gmail: [%]
• Apple Mail: [%]
• Outlook: [%]
• Other: [%]

Geographic Performance:
[If relevant - which locations had best engagement]

POST-CAMPAIGN ANALYSIS:

What Worked Well:
• [Success 1 - with specific data]
• [Success 2 - with specific data]
• [Success 3 - with specific data]

What Didn't Work:
• [Challenge 1 - with specific data]
• [Challenge 2 - with specific data]

Key Learnings:
• [Learning 1]
• [Learning 2]
• [Learning 3]

Optimization Recommendations:
• [Recommendation for future campaigns]
• [Recommendation for future campaigns]
• [Recommendation for future campaigns]

FOLLOW-UP ACTIONS:

☐ Segment engaged users (openers/clickers) for next campaign
☐ Create re-engagement sequence for non-openers
☐ Remove hard bounces from list
☐ Update email preferences based on feedback
☐ Document learnings in campaign playbook
☐ Share results with team
☐ Plan follow-up campaign based on insights
☐ Thank customers who converted
☐ Survey non-converters for feedback

APPROVALS:

Campaign Strategy:
Approved by: _________________ Date: _______
[Marketing Manager]

Email Copy:
Approved by: _________________ Date: _______
[Content Lead]

Email Design:
Approved by: _________________ Date: _______
[Design Lead]

Legal/Compliance:
Approved by: _________________ Date: _______
[Legal/Compliance Officer]

Final Launch Approval:
Approved by: _________________ Date: _______
[VP Marketing/CMO]

NOTES & ADDITIONAL DETAILS:
[Space for any additional campaign notes, special considerations, or context]
    `
  }
];

// Add marketing templates to the main templates array
templates.push(...marketingTemplates);

class DocumentService {
  getTemplates(): DocumentTemplate[] {
    return templates;
  }

  async generateDocx(template: DocumentTemplate, variables: Record<string, string> = {}): Promise<void> {
    try {
      const doc = new Document({
        sections: [{
          properties: {},
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: template.title,
                  bold: true,
                  size: 32,
                }),
              ],
              heading: HeadingLevel.TITLE,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: " ",
                  size: 24,
                }),
              ],
            }),
            ...this.parseContentToParagraphs(template.content || ''),
          ],
        }],
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `${template.title}.docx`);
    } catch (error) {
      console.error('Error generating document:', error);
      throw error;
    }
  }

  private parseContentToParagraphs(content: string): Paragraph[] {
    const lines = content.split('\n');
    const paragraphs: Paragraph[] = [];

    for (const line of lines) {
      const trimmedLine = line.trim();
      
      if (trimmedLine === '') {
        paragraphs.push(new Paragraph({
          children: [new TextRun({ text: " " })],
        }));
        continue;
      }

      // Handle headers (lines that are all caps or start with numbers)
      if (trimmedLine.match(/^[A-Z\s&()]+:?$/) || trimmedLine.match(/^\d+\.\s+[A-Z]/)) {
        paragraphs.push(new Paragraph({
          children: [
            new TextRun({
              text: trimmedLine,
              bold: true,
              size: 24,
            }),
          ],
          spacing: { before: 200, after: 100 },
        }));
        continue;
      }

      // Handle bullet points
      if (trimmedLine.startsWith('•') || trimmedLine.startsWith('-')) {
        paragraphs.push(new Paragraph({
          children: [
            new TextRun({
              text: trimmedLine,
              size: 22,
            }),
          ],
          bullet: { level: 0 },
          spacing: { before: 50, after: 50 },
        }));
        continue;
      }

      // Handle checkboxes
      if (trimmedLine.startsWith('☐')) {
        paragraphs.push(new Paragraph({
          children: [
            new TextRun({
              text: trimmedLine,
              size: 22,
            }),
          ],
          spacing: { before: 50, after: 50 },
        }));
        continue;
      }

      // Regular paragraph
      paragraphs.push(new Paragraph({
        children: [
          new TextRun({
            text: trimmedLine,
            size: 22,
          }),
        ],
        spacing: { before: 100, after: 100 },
      }));
    }

    return paragraphs;
  }

  getTemplatesByCategory(category: string): DocumentTemplate[] {
    return templates.filter(template => template.category === category);
  }

  getAllCategories(): string[] {
    const categories = [...new Set(templates.map(template => template.category))];
    return categories.sort();
  }

  searchTemplates(query: string): DocumentTemplate[] {
    const lowercaseQuery = query.toLowerCase();
    return templates.filter(template =>
      template.title.toLowerCase().includes(lowercaseQuery) ||
      template.description.toLowerCase().includes(lowercaseQuery) ||
      template.category.toLowerCase().includes(lowercaseQuery) ||
      template.tags.some(tag => tag.toLowerCase().includes(lowercaseQuery))
    );
  }
}

export const documentService = new DocumentService();