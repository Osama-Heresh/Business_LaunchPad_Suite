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
  },
  {
    id: 'marketing-landing-page-copy',
    title: 'Landing Page Copy Template',
    description: 'High-converting landing page copywriting framework with all essential conversion elements',
    category: 'marketing',
    downloadCount: 6120,
    fileSize: '41.5 KB',
    rating: 4.8,
    tags: ['landing-page', 'copywriting', 'conversion', 'sales-page', 'marketing'],
    createdAt: '2024-01-13T11:00:00Z',
    updatedAt: '2024-02-02T14:15:00Z',
    content: `
LANDING PAGE COPY TEMPLATE

PAGE INFORMATION:
Page Name: [Landing Page Name]
URL Slug: /[page-url-slug]
Campaign: [Campaign Name]
Primary Goal: [Main conversion goal - e.g., Sign ups, Sales, Downloads]
Target Audience: [Primary audience description]
Created By: [Name]
Date: [Date]
Status: ☐ Draft ☐ Review ☐ Design ☐ Development ☐ Testing ☐ Live

TARGET AUDIENCE PROFILE:

Primary Persona: [Persona name]
• Age: [Range]
• Occupation: [Job title/role]
• Income: [Range]
• Location: [Geographic area]

Pain Points:
• [Pain point 1]
• [Pain point 2]
• [Pain point 3]

Desires/Goals:
• [Goal 1]
• [Goal 2]
• [Goal 3]

Objections to Address:
• [Objection 1]
• [Objection 2]
• [Objection 3]

UNIQUE VALUE PROPOSITION:
[One clear sentence explaining what you offer, who it's for, and why it's different/better]

COMPETITIVE ADVANTAGE:
What makes us different:
• [Differentiator 1]
• [Differentiator 2]
• [Differentiator 3]

---

SECTION 1: HERO SECTION (Above the Fold)

Main Headline (Focus on benefit, not feature):
[Powerful headline that addresses main pain point or desire - 8-12 words]

Subheadline (Expand on the headline):
[Elaborate on the headline, add specificity or urgency - 15-25 words]

Hero Visual:
☐ Hero image showing result/benefit
☐ Product screenshot/mockup
☐ Explainer video (Length: [X] seconds/minutes)
☐ Background image with text overlay
☐ Illustration/graphic

Alt Text: [SEO-optimized description]

Primary CTA Button:
Button Text: [Action verb + benefit, e.g., "Start Your Free Trial"]
Button Color: [High-contrast color]
Button Size: [Large, prominent]
Link/Action: [URL or form trigger]

Trust Elements (Choose 2-3):
☐ Customer count: "Join [X,XXX]+ happy customers"
☐ Social proof: "[X] companies trust us"
☐ Media logos: Featured in [Publication 1, Publication 2]
☐ Awards/Badges: [Award name, certification]
☐ Security badges: [SSL, payment processor logos]
☐ Guarantee: "30-Day Money-Back Guarantee"
☐ Rating: "★★★★★ [X.X] stars on [Platform]"

---

SECTION 2: PROBLEM/AGITATION

Section Headline:
[Question or statement that identifies the problem - make it relatable]

Subheadline (Optional):
[Reinforce the problem]

Body Copy:

Paragraph 1 - Identify the Problem:
[Describe the specific problem your audience faces. Make it relatable and emotional.]

Paragraph 2 - Agitate the Problem:
[Explain what happens if they don't solve it. Paint a picture of the consequences.]

Paragraph 3 - Hint at Solution:
[Transition to hope - there IS a solution]

Problem List (Optional):
Are you struggling with:
• [Specific problem 1]
• [Specific problem 2]
• [Specific problem 3]
• [Specific problem 4]

Visual Element:
☐ Image showing the problem
☐ Before/After comparison
☐ Pain point infographic
☐ Relatable scenario illustration

---

SECTION 3: SOLUTION INTRODUCTION

Section Headline:
[Introduce your solution as the answer to their problem]

Subheadline:
[How your solution makes their life better]

Body Copy:
[2-3 paragraphs explaining what your product/service is and how it solves the problem. Focus on transformation, not just features.]

Solution Overview (3-4 Key Benefits):
• [Benefit 1]: [How it helps them]
• [Benefit 2]: [How it helps them]
• [Benefit 3]: [How it helps them]
• [Benefit 4]: [How it helps them]

Visual Element:
☐ Product image/screenshot
☐ Demo video
☐ Animated GIF showing product in action
☐ Diagram of how it works

CTA:
[Repeat primary CTA or softer CTA like "Learn More"]

---

SECTION 4: HOW IT WORKS

Section Headline:
[Simple headline - e.g., "How It Works" or "Get Started in 3 Easy Steps"]

Step 1:
Icon: [Icon description]
Title: [Step title]
Description: [Brief explanation - 1-2 sentences]

Step 2:
Icon: [Icon description]
Title: [Step title]
Description: [Brief explanation - 1-2 sentences]

Step 3:
Icon: [Icon description]
Title: [Step title]
Description: [Brief explanation - 1-2 sentences]

Optional Step 4:
Icon: [Icon description]
Title: [Step title]
Description: [Brief explanation - 1-2 sentences]

Closing Copy:
[Emphasize how easy/fast it is]

CTA:
Button Text: [CTA text]

---

SECTION 5: FEATURES & BENEFITS

Section Headline:
[Headline focusing on benefits, not just features]

Feature 1:
Icon/Image: [Visual element]
Feature Name: [Name]
Description: [2-3 sentences focusing on the BENEFIT to the user]

Feature 2:
Icon/Image: [Visual element]
Feature Name: [Name]
Description: [2-3 sentences focusing on the BENEFIT to the user]

Feature 3:
Icon/Image: [Visual element]
Feature Name: [Name]
Description: [2-3 sentences focusing on the BENEFIT to the user]

Feature 4:
Icon/Image: [Visual element]
Feature Name: [Name]
Description: [2-3 sentences focusing on the BENEFIT to the user]

Feature 5:
Icon/Image: [Visual element]
Feature Name: [Name]
Description: [2-3 sentences focusing on the BENEFIT to the user]

Feature 6:
Icon/Image: [Visual element]
Feature Name: [Name]
Description: [2-3 sentences focusing on the BENEFIT to the user]

---

SECTION 6: SOCIAL PROOF & TESTIMONIALS

Section Headline:
[e.g., "Join [X,XXX] Happy Customers" or "What Our Customers Are Saying"]

Statistics Bar (Optional):
• [X,XXX] Customers Served
• [XX%] Satisfaction Rate
• ★ [X.X] Average Rating
• [X] Years in Business

Testimonial 1:
Quote: "[Specific result or transformation - be concrete with numbers if possible]"
Customer Name: [Full Name]
Title/Company: [Title at Company]
Photo: [Customer headshot]
Result Highlight: [Specific metric, e.g., "Increased sales by 150%"]

Testimonial 2:
Quote: "[Testimonial highlighting different benefit]"
Customer Name: [Full Name]
Title/Company: [Title at Company]
Photo: [Customer headshot]
Result Highlight: [Specific metric]

Testimonial 3:
Quote: "[Testimonial addressing common objection]"
Customer Name: [Full Name]
Title/Company: [Title at Company]
Photo: [Customer headshot]
Result Highlight: [Specific metric]

Video Testimonial (Optional):
Customer: [Name]
Video Length: [Duration]
Key Quote: "[Most powerful quote from video]"

Case Study Highlight (Optional):
Company: [Company name]
Challenge: [What problem they had]
Solution: [How you helped]
Results:
• [Metric 1]: [Improvement]
• [Metric 2]: [Improvement]
• [Metric 3]: [Improvement]

CTA: "Read Full Case Study →"

Review Platform Highlights:
[Platform 1]: ★★★★★ [X.X] stars ([XXX] reviews)
[Platform 2]: ★★★★★ [X.X] stars ([XXX] reviews)
[Platform 3]: ★★★★★ [X.X] stars ([XXX] reviews)

---

SECTION 7: TRUST & CREDIBILITY

Section Headline (Optional):
[e.g., "Trusted by Industry Leaders" or "Recognized Excellence"]

Client Logos:
[Logo 1] [Logo 2] [Logo 3] [Logo 4] [Logo 5]
[Logo 6] [Logo 7] [Logo 8] [Logo 9] [Logo 10]

Press Mentions:
"[Headline of feature]" - [Publication Name]
"[Headline of feature]" - [Publication Name]
"[Headline of feature]" - [Publication Name]

Certifications & Awards:
• [Certification/Award 1]
• [Certification/Award 2]
• [Certification/Award 3]

Security & Compliance:
☐ SSL Secured
☐ GDPR Compliant
☐ SOC 2 Certified
☐ PCI Compliant
☐ [Industry-specific certification]

---

SECTION 8: PRICING (If applicable)

Section Headline:
[Pricing headline that frames value]

Pricing Strategy:
☐ Single price point
☐ Tiered pricing (2-3 tiers)
☐ Custom pricing
☐ Freemium model

TIER 1: [Plan Name]
Price: $[Amount]/[month/year]
Billing: [Monthly/Annually]
Best For: [Who this is for]
Features Included:
• [Feature 1]
• [Feature 2]
• [Feature 3]
• [Feature 4]
• [Feature 5]

CTA Button: [Button text]

TIER 2: [Plan Name] ⭐ MOST POPULAR
Price: $[Amount]/[month/year]
Billing: [Monthly/Annually]
Savings: [Save X% vs monthly]
Best For: [Who this is for]
Features Included:
• Everything in [Tier 1]
• [Additional feature 1]
• [Additional feature 2]
• [Additional feature 3]
• [Additional feature 4]
• [Additional feature 5]

CTA Button: [Button text]
Badge: "Most Popular" or "Best Value"

TIER 3: [Plan Name]
Price: $[Amount]/[month/year] or "Custom"
Billing: [Monthly/Annually/Custom]
Best For: [Who this is for]
Features Included:
• Everything in [Tier 2]
• [Premium feature 1]
• [Premium feature 2]
• [Premium feature 3]
• [Premium feature 4]
• Dedicated account manager

CTA Button: [Button text]

Pricing Notes:
• No credit card required
• Cancel anytime
• 30-day money-back guarantee
• All plans include [core feature]

Comparison Table (Optional):
Feature | [Tier 1] | [Tier 2] | [Tier 3]
--------|----------|----------|----------
[Feature 1] | ✓ | ✓ | ✓
[Feature 2] | ✓ | ✓ | ✓
[Feature 3] | - | ✓ | ✓
[Feature 4] | - | ✓ | ✓
[Feature 5] | - | - | ✓

---

SECTION 9: OBJECTION HANDLING (FAQ)

Section Headline:
[e.g., "Frequently Asked Questions" or "Got Questions? We've Got Answers"]

Question 1: [Most common question or objection]
Answer: [Clear, reassuring answer that addresses concern]

Question 2: [Question about pricing/value]
Answer: [Explain value, ROI, or payment options]

Question 3: [Question about ease of use]
Answer: [Explain how easy it is, mention support]

Question 4: [Question about results/timeline]
Answer: [Set realistic expectations, provide examples]

Question 5: [Question about guarantee/risk]
Answer: [Explain guarantee, remove risk]

Question 6: [Technical question]
Answer: [Simple, non-technical explanation]

Question 7: [Comparison question]
Answer: [Differentiate from competitors]

Question 8: [Implementation question]
Answer: [Explain onboarding/setup process]

Still have questions?
[Link to full FAQ page or "Contact Us" button]

---

SECTION 10: FINAL CTA (Last Chance)

Section Headline:
[Benefit-focused headline with urgency - e.g., "Ready to Transform Your [Area]?"]

Subheadline:
[Reinforce value or add scarcity element]

Body Copy (Optional):
[Brief recap of key benefits or special offer]

Primary CTA Button:
Button Text: [Strong action verb + benefit]
Button Color: [High contrast]
Button Size: [Extra large]

Risk Reversal Elements:
• ✓ 30-Day Money-Back Guarantee
• ✓ No Credit Card Required
• ✓ Cancel Anytime
• ✓ Free [Trial/Consultation/Assessment]

Trust Signals:
[Payment badges: Visa, Mastercard, PayPal, etc.]
[Security badge: SSL, Norton, etc.]
"Join [X,XXX]+ satisfied customers"

Urgency Element (Optional, use carefully):
☐ Limited-time offer: "Offer ends [Date]"
☐ Limited quantity: "Only [X] spots available"
☐ Bonus for quick action: "Sign up today and get [Bonus]"
☐ Countdown timer

---

FOOTER

Company Information:
[Company Name]
[Address]
[Phone]
[Email]

Navigation Links:
• About Us
• Contact
• Privacy Policy
• Terms of Service
• Refund Policy
• Support/Help Center
• Blog

Social Media:
[Facebook] [Twitter] [LinkedIn] [Instagram] [YouTube]

Additional Links:
• Affiliate Program
• Careers
• Press/Media Kit
• Sitemap

Copyright:
© [Year] [Company Name]. All rights reserved.

---

CONVERSION OPTIMIZATION ELEMENTS:

Exit-Intent Popup:
Trigger: Mouse moves toward close/back button
Headline: [Urgency headline, e.g., "Wait! Don't Miss Out!"]
Offer: [Special discount, bonus, or lead magnet]
CTA: [Button text]

Floating/Sticky CTA Bar:
Appears After: [Scrolling past hero section]
Message: [Brief value prop]
CTA Button: [Text]
Position: [Top/Bottom]

Chat Widget:
Position: Bottom right
Initial Message: [Greeting]
Trigger: [After X seconds or X% scroll]

Countdown Timer (If applicable):
Location: [Near main CTA]
Message: "Special offer ends in:"
Duration: [Time period]

Lead Magnet/Content Upgrade:
Offer: [Free resource, checklist, guide, etc.]
Headline: [Benefit of the free resource]
Form Fields: [Name, Email only - keep it simple]

---

COPYWRITING GUIDELINES CHECKLIST:

Headline & Copy:
☐ Headline addresses primary pain point or desire
☐ Benefits emphasized over features throughout
☐ Clear, specific value proposition above fold
☐ Emotional + logical appeal balanced
☐ Active voice used (avoid passive)
☐ Second person "you" used frequently
☐ Specific numbers and data included
☐ No jargon - written at 8th grade reading level
☐ Short sentences and paragraphs (3-4 lines max)
☐ Transition words between sections

Conversion Elements:
☐ Single, clear conversion goal
☐ Multiple CTAs throughout page (every 1-2 screens)
☐ CTA copy is action-oriented and benefit-focused
☐ Risk reversal elements included
☐ Social proof throughout
☐ Objections addressed in FAQ
☐ Urgency or scarcity element (if authentic)
☐ Clear next steps explained

Visual Hierarchy:
☐ Scannable format (bullets, subheads, whitespace)
☐ Important elements stand out visually
☐ Eye flow guides to CTA
☐ F-pattern or Z-pattern layout

---

SEO & TECHNICAL ELEMENTS:

Page Title Tag: [60 characters max, include primary keyword]
Meta Description: [160 characters max, compelling with keyword and CTA]
H1 Tag: [Main headline - only one H1 per page]
H2 Tags: [Section headlines]
Image Alt Text: [All images have descriptive alt text]
URL Structure: /[keyword-focused-slug]

Schema Markup:
☐ Product schema (if applicable)
☐ Review schema (for testimonials)
☐ FAQ schema
☐ Organization schema

Mobile Optimization:
☐ Mobile-first responsive design
☐ Touch-friendly buttons (44x44px min)
☐ Readable font sizes (16px+ for body)
☐ No horizontal scrolling
☐ Fast load time (<3 seconds)
☐ Compressed images

Page Speed:
Target: <3 seconds load time
☐ Images optimized and compressed
☐ CSS/JS minified
☐ Browser caching enabled
☐ CDN enabled
☐ Lazy loading for images

Browser Compatibility:
☐ Chrome
☐ Safari
☐ Firefox
☐ Edge
☐ Mobile browsers (iOS Safari, Chrome Mobile)

---

A/B TESTING IDEAS:

Headlines to Test:
• [Variation 1]
• [Variation 2]
• [Variation 3]

CTA Button Copy:
• [Variation 1]
• [Variation 2]

CTA Button Color:
• [Color 1]
• [Color 2]

Hero Section:
• Image vs. Video
• Short form vs. Long form
• Different value propositions

Pricing Display:
• Monthly vs. Annual pricing shown first
• 2 tiers vs. 3 tiers
• Pricing table vs. Cards

Social Proof Placement:
• Above fold vs. Below features
• Video testimonials vs. Text

Form Length:
• 2 fields vs. 3 fields vs. 4 fields
• One-step vs. Multi-step

---

TRACKING & ANALYTICS:

Conversion Goals:
Primary Goal: [Main conversion - purchase, signup, etc.]
Micro-Conversions:
• Video play
• Scroll to pricing
• FAQ expansion
• Chat initiation
• Email signup

Events to Track:
• CTA button clicks (all instances)
• Form submissions (started vs. completed)
• Video plays and completion rate
• Scroll depth (25%, 50%, 75%, 100%)
• Time on page
• Exit intent popup triggers
• Live chat starts
• Outbound link clicks
• Phone number clicks (mobile)

Heatmap Tools:
☐ Hotjar
☐ Crazy Egg
☐ Microsoft Clarity

Analytics Platform:
☐ Google Analytics 4
☐ Facebook Pixel
☐ LinkedIn Insight Tag
☐ Google Tag Manager

---

LAUNCH CHECKLIST:

Pre-Launch:
☐ All copy proofread and approved
☐ Legal review completed (if needed)
☐ All links tested
☐ Forms tested and working
☐ Payment processor tested (if applicable)
☐ Mobile responsive design verified
☐ Page speed optimized
☐ Browser compatibility checked
☐ Analytics and tracking installed
☐ Conversion goals set up
☐ Privacy policy linked
☐ Terms of service linked
☐ SSL certificate active
☐ 404 error page set up
☐ Thank you page created
☐ Confirmation emails set up

Post-Launch:
☐ Monitor analytics daily
☐ Track conversion rate
☐ Review heatmaps weekly
☐ Collect user feedback
☐ A/B test elements
☐ Iterate based on data
☐ Monitor page speed
☐ Check for broken links monthly
☐ Update testimonials regularly
☐ Refresh copy quarterly

---

PERFORMANCE TARGETS:

Conversion Rate: [Target %]
Average Time on Page: [Target minutes]
Bounce Rate: [Target %]
Form Completion Rate: [Target %]
Cost Per Acquisition: $[Target amount]
Return on Ad Spend: [Target multiplier]

Review Schedule:
• Daily: Monitor traffic and conversions
• Weekly: Review heatmaps and user behavior
• Monthly: Comprehensive performance analysis
• Quarterly: Major updates and optimization

---

APPROVALS:

Copy Written by: _________________ Date: _______
[Copywriter Name]

Copy Reviewed by: _________________ Date: _______
[Marketing Manager]

Design Approved by: _________________ Date: _______
[Design Lead]

Legal Approved by: _________________ Date: _______
[Legal Team]

Final Launch Approval: _________________ Date: _______
[Executive/CMO]

NOTES:
[Additional notes, special considerations, or context for this landing page]
    `
  },
  {
    id: 'marketing-plan-one-pager',
    title: 'Marketing Plan (One-Pager)',
    description: 'Concise one-page marketing plan covering strategy, tactics, budget, and key metrics',
    category: 'marketing',
    downloadCount: 4950,
    fileSize: '34.8 KB',
    rating: 4.7,
    tags: ['marketing-plan', 'strategy', 'planning', 'one-pager', 'campaign'],
    createdAt: '2024-01-15T14:20:00Z',
    updatedAt: '2024-02-03T10:45:00Z',
    content: `
MARKETING PLAN (ONE-PAGER)

COMPANY & CAMPAIGN INFO:
Company: [Company Name]
Product/Service: [What you're marketing]
Time Period: [Q1 2024, FY 2024, etc.]
Plan Owner: [Marketing Manager Name]
Date Created: [Date]
Total Budget: $[Amount]

---

BUSINESS OBJECTIVES (What the company wants to achieve):

1. [Primary business goal - e.g., Increase revenue by 25%]
2. [Secondary goal - e.g., Expand into new market segment]
3. [Tertiary goal - e.g., Improve customer retention by 15%]

---

MARKETING OBJECTIVES (SMART Goals):

1. [Specific goal with number and deadline]
   Example: Generate 500 qualified leads per month by end of Q2

2. [Specific goal with number and deadline]
   Example: Increase website traffic to 50,000 monthly visitors by Q4

3. [Specific goal with number and deadline]
   Example: Achieve 10% conversion rate on landing pages by end of Q3

---

TARGET AUDIENCE:

Primary Persona: "[Persona Name]"
• Demographics: [Age range, location, income, job title]
• Psychographics: [Interests, values, lifestyle]
• Pain Points: [Top 3 challenges they face]
• Goals: [What they want to achieve]
• Where to Reach Them: [Channels/platforms they use]

Secondary Persona: "[Persona Name]"
• Demographics: [Age range, location, income, job title]
• Psychographics: [Interests, values, lifestyle]
• Pain Points: [Top 3 challenges they face]

---

VALUE PROPOSITION:

[One clear sentence explaining what you offer, who it's for, and why you're the best choice]

Example: "We help [target audience] achieve [desired outcome] through [unique approach/solution] without [common pain point]."

---

KEY MESSAGES (Consistent across all channels):

1. [Core message about main benefit]
2. [Core message about differentiation]
3. [Core message about credibility/trust]

---

MARKETING STRATEGIES & TACTICS:

1. DIGITAL MARKETING

Content Marketing:
• Blog posts: [X] per week on [topics]
• Ebooks/Guides: [X] per quarter
• Videos: [X] per month ([platform])
• Infographics: [X] per quarter

Social Media:
• Platforms: [LinkedIn, Instagram, Twitter, etc.]
• Posting frequency: [X] times per week per platform
• Paid social budget: $[amount]/month
• Focus: [Brand awareness / Lead generation / Community building]

Email Marketing:
• Newsletter: [Frequency]
• Drip campaigns: [Number of sequences]
• List growth goal: [X] new subscribers/month
• Target open rate: [%]

SEO & SEM:
• Target keywords: [Top 5 keywords]
• Content optimization: [X] pages/month
• Google Ads budget: $[amount]/month
• Target: [X] organic visitors/month

Paid Advertising:
• Google Ads: $[amount]/month
• Facebook/Instagram Ads: $[amount]/month
• LinkedIn Ads: $[amount]/month
• Retargeting: $[amount]/month

Website & Landing Pages:
• New landing pages: [X] per quarter
• A/B tests: [What to test]
• Conversion rate goal: [%]

2. TRADITIONAL MARKETING (If applicable)

☐ Trade Shows/Events: [Which ones, budget: $X]
☐ Print Advertising: [Where, budget: $X]
☐ Direct Mail: [Frequency, budget: $X]
☐ Radio/TV: [Channels, budget: $X]

3. PUBLIC RELATIONS & PARTNERSHIPS

☐ Press releases: [X] per quarter
☐ Media outreach: [Target publications]
☐ Speaking engagements: [Events/conferences]
☐ Strategic partnerships: [Potential partners]
☐ Influencer collaborations: [Number, budget: $X]

---

CONTENT CALENDAR (High-Level):

Q1: [Theme/Focus]
• Key campaigns: [Campaign 1, Campaign 2]
• Major content: [Ebook, webinar, etc.]

Q2: [Theme/Focus]
• Key campaigns: [Campaign 1, Campaign 2]
• Major content: [Video series, whitepaper, etc.]

Q3: [Theme/Focus]
• Key campaigns: [Campaign 1, Campaign 2]
• Major content: [Case studies, event, etc.]

Q4: [Theme/Focus]
• Key campaigns: [Campaign 1, Campaign 2]
• Major content: [Year-end report, holiday campaign, etc.]

---

BUDGET BREAKDOWN:

Total Marketing Budget: $[Amount]

Allocation by Category:
• Digital Advertising: $[Amount] ([%]%)
• Content Creation: $[Amount] ([%]%)
• Marketing Tools/Software: $[Amount] ([%]%)
• Events/Trade Shows: $[Amount] ([%]%)
• PR & Media: $[Amount] ([%]%)
• Email Marketing Platform: $[Amount] ([%]%)
• Social Media Management: $[Amount] ([%]%)
• Design/Creative Services: $[Amount] ([%]%)
• Analytics/Research: $[Amount] ([%]%)
• Contingency (10%): $[Amount] ([%]%)

Monthly Budget: $[Total ÷ 12]

---

KEY PERFORMANCE INDICATORS (KPIs):

Lead Generation:
• Monthly leads: [Target number]
• Marketing Qualified Leads (MQLs): [Target number]
• Lead-to-customer rate: [%]
• Cost per lead: $[Amount]

Website & Traffic:
• Monthly website visitors: [Target]
• Conversion rate: [%]
• Bounce rate: [Target %]
• Average session duration: [Minutes]

Social Media:
• Total followers: [Target across all platforms]
• Engagement rate: [%]
• Social referral traffic: [Target]

Email Marketing:
• List size: [Target subscribers]
• Open rate: [%]
• Click-through rate: [%]
• Conversion rate: [%]

Content Performance:
• Blog traffic: [Monthly visitors]
• Video views: [Total/month]
• Content downloads: [Number/month]

Revenue & ROI:
• Marketing-attributed revenue: $[Amount]
• Customer Acquisition Cost (CAC): $[Amount]
• Return on Marketing Investment (ROMI): [X:1]
• Customer Lifetime Value (CLV): $[Amount]

Brand Awareness:
• Brand search volume: [Target searches/month]
• Share of voice: [%]
• Media mentions: [Number/quarter]
• Social sentiment: [Positive %]

---

COMPETITIVE ANALYSIS:

Top 3 Competitors:

1. [Competitor 1]
   Strengths: [What they do well]
   Weaknesses: [Where they fall short]
   Our Advantage: [How we're better]

2. [Competitor 2]
   Strengths: [What they do well]
   Weaknesses: [Where they fall short]
   Our Advantage: [How we're better]

3. [Competitor 3]
   Strengths: [What they do well]
   Weaknesses: [Where they fall short]
   Our Advantage: [How we're better]

Market Opportunities:
• [Untapped opportunity 1]
• [Untapped opportunity 2]

Market Threats:
• [Potential threat 1]
• [Potential threat 2]

---

TIMELINE & MAJOR MILESTONES:

Month 1-3 (Q1):
• [Major activity/campaign]
• [Major activity/campaign]
• Milestone: [Deliverable/Achievement]

Month 4-6 (Q2):
• [Major activity/campaign]
• [Major activity/campaign]
• Milestone: [Deliverable/Achievement]

Month 7-9 (Q3):
• [Major activity/campaign]
• [Major activity/campaign]
• Milestone: [Deliverable/Achievement]

Month 10-12 (Q4):
• [Major activity/campaign]
• [Major activity/campaign]
• Milestone: [Deliverable/Achievement]

---

TEAM & RESOURCES:

Internal Team:
• [Name/Role]: [Primary responsibilities]
• [Name/Role]: [Primary responsibilities]
• [Name/Role]: [Primary responsibilities]

External Resources:
• Agency: [Name] - [Services provided]
• Freelancers: [Roles - copywriter, designer, etc.]
• Consultants: [Specialty]

Marketing Technology Stack:
• CRM: [Platform]
• Email Marketing: [Platform]
• Analytics: [Platform]
• Social Media Management: [Platform]
• Content Management: [Platform]
• Marketing Automation: [Platform]
• Design Tools: [Platform]

---

RISKS & MITIGATION STRATEGIES:

Risk 1: [Potential problem]
Likelihood: ☐ High ☐ Medium ☐ Low
Impact: ☐ High ☐ Medium ☐ Low
Mitigation: [How to prevent or address]

Risk 2: [Potential problem]
Likelihood: ☐ High ☐ Medium ☐ Low
Impact: ☐ High ☐ Medium ☐ Low
Mitigation: [How to prevent or address]

Risk 3: [Potential problem]
Likelihood: ☐ High ☐ Medium ☐ Low
Impact: ☐ High ☐ Medium ☐ Low
Mitigation: [How to prevent or address]

---

MEASUREMENT & REPORTING:

Reporting Frequency:
• Weekly: Quick metrics check (traffic, leads, spend)
• Monthly: Comprehensive performance review
• Quarterly: Strategic review and optimization

Dashboard Tools:
• [Google Analytics, HubSpot, etc.]
• Custom dashboard: [Platform]

Monthly Report Includes:
• KPI performance vs. targets
• Campaign results
• Budget utilization
• Key wins and challenges
• Recommendations for next month

Review Meetings:
• Weekly team sync: [Day/Time]
• Monthly stakeholder review: [Day/Time]
• Quarterly strategic planning: [When]

---

SUCCESS CRITERIA (How we'll know we succeeded):

☐ Achieved [X]% of lead generation goal
☐ Met or exceeded revenue target of $[Amount]
☐ Maintained CAC under $[Amount]
☐ Achieved ROMI of at least [X:1]
☐ Increased brand awareness by [X]%
☐ Launched [X] successful campaigns
☐ Grew email list to [X] subscribers
☐ Achieved [X]% customer satisfaction score

---

OPTIMIZATION APPROACH:

Continuous Improvement:
• Weekly: Adjust ad spend based on performance
• Monthly: Optimize underperforming campaigns
• Quarterly: Major strategy pivots if needed

A/B Testing Focus:
• Landing page headlines and CTAs
• Email subject lines and send times
• Ad creative and copy
• Social media post formats

---

APPROVALS:

Marketing Plan Prepared by: _________________ Date: _______
[Marketing Manager Name]

Reviewed by: _________________ Date: _______
[Department Head/Director]

Budget Approved by: _________________ Date: _______
[Finance/CFO]

Final Approval by: _________________ Date: _______
[CEO/Executive Team]

---

NOTES & STRATEGIC CONTEXT:
[Additional context, assumptions, or important considerations for this plan]
    `
  },
  {
    id: 'marketing-swot-analysis',
    title: 'SWOT Analysis Template',
    description: 'Comprehensive SWOT analysis framework for evaluating strengths, weaknesses, opportunities, and threats',
    category: 'marketing',
    downloadCount: 5670,
    fileSize: '38.4 KB',
    rating: 4.6,
    tags: ['swot', 'analysis', 'strategy', 'planning', 'competitive'],
    createdAt: '2024-01-17T09:45:00Z',
    updatedAt: '2024-02-04T13:20:00Z',
    content: `
SWOT ANALYSIS

ANALYSIS INFORMATION:
Company/Product: [Name]
Industry: [Industry]
Analysis Date: [Date]
Prepared By: [Name and Title]
Purpose: [Why this analysis is being conducted]
Time Period Analyzed: [Past period considered]
Future Outlook Period: [Future period projected]

EXECUTIVE SUMMARY:
[2-3 sentence overview of key findings and strategic implications]

---

INTERNAL FACTORS

STRENGTHS (What we do well | What advantages do we have?)

1. [Strength 1]
   Category: ☐ Brand ☐ Product ☐ Financial ☐ Operational ☐ People ☐ Technology
   Description: [Detailed explanation of this strength]
   Evidence: [Data, metrics, or examples that prove this]
   Competitive Impact: [How this gives us an edge]
   Importance: ☐ Critical ☐ High ☐ Medium ☐ Low

2. [Strength 2]
   Category: ☐ Brand ☐ Product ☐ Financial ☐ Operational ☐ People ☐ Technology
   Description: [Detailed explanation]
   Evidence: [Supporting data]
   Competitive Impact: [Advantage created]
   Importance: ☐ Critical ☐ High ☐ Medium ☐ Low

3. [Strength 3]
   Category: ☐ Brand ☐ Product ☐ Financial ☐ Operational ☐ People ☐ Technology
   Description: [Detailed explanation]
   Evidence: [Supporting data]
   Competitive Impact: [Advantage created]
   Importance: ☐ Critical ☐ High ☐ Medium ☐ Low

[Continue for 5-10 total strengths]

STRENGTHS BY CATEGORY:

Brand & Market Position:
• [Specific strength related to brand]
• [Specific strength related to reputation]
• [Specific strength related to market share]

Products/Services:
• [Specific strength related to offerings]
• [Specific strength related to quality]
• [Specific strength related to innovation]

Financial Resources:
• [Specific strength related to capital]
• [Specific strength related to profitability]
• [Specific strength related to cash flow]

Operations & Processes:
• [Specific strength related to efficiency]
• [Specific strength related to systems]
• [Specific strength related to scale]

People & Culture:
• [Specific strength related to team]
• [Specific strength related to expertise]
• [Specific strength related to culture]

Technology & Innovation:
• [Specific strength related to tech stack]
• [Specific strength related to IP]
• [Specific strength related to R&D]

Customer Relationships:
• [Specific strength related to loyalty]
• [Specific strength related to retention]
• [Specific strength related to satisfaction]

Distribution & Partnerships:
• [Specific strength related to channels]
• [Specific strength related to partners]
• [Specific strength related to reach]

---

WEAKNESSES (What could we improve? | Where do we fall short?)

1. [Weakness 1]
   Category: ☐ Brand ☐ Product ☐ Financial ☐ Operational ☐ People ☐ Technology
   Description: [Detailed explanation of this weakness]
   Evidence: [Data showing the gap or problem]
   Impact: [How this hurts us or limits growth]
   Root Cause: [Why this weakness exists]
   Severity: ☐ Critical ☐ High ☐ Medium ☐ Low
   Priority to Address: ☐ Immediate ☐ Short-term ☐ Medium-term ☐ Low priority

2. [Weakness 2]
   Category: ☐ Brand ☐ Product ☐ Financial ☐ Operational ☐ People ☐ Technology
   Description: [Detailed explanation]
   Evidence: [Supporting data]
   Impact: [Consequences]
   Root Cause: [Why it exists]
   Severity: ☐ Critical ☐ High ☐ Medium ☐ Low
   Priority to Address: ☐ Immediate ☐ Short-term ☐ Medium-term ☐ Low priority

3. [Weakness 3]
   Category: ☐ Brand ☐ Product ☐ Financial ☐ Operational ☐ People ☐ Technology
   Description: [Detailed explanation]
   Evidence: [Supporting data]
   Impact: [Consequences]
   Root Cause: [Why it exists]
   Severity: ☐ Critical ☐ High ☐ Medium ☐ Low
   Priority to Address: ☐ Immediate ☐ Short-term ☐ Medium-term ☐ Low priority

[Continue for 5-10 total weaknesses]

WEAKNESSES BY CATEGORY:

Brand & Market Position:
• [Specific weakness]
• [Specific weakness]

Products/Services:
• [Specific weakness]
• [Specific weakness]

Financial Resources:
• [Specific weakness]
• [Specific weakness]

Operations & Processes:
• [Specific weakness]
• [Specific weakness]

People & Culture:
• [Specific weakness]
• [Specific weakness]

Technology & Innovation:
• [Specific weakness]
• [Specific weakness]

Customer Relationships:
• [Specific weakness]
• [Specific weakness]

Distribution & Partnerships:
• [Specific weakness]
• [Specific weakness]

---

EXTERNAL FACTORS

OPPORTUNITIES (What favorable external conditions exist?)

1. [Opportunity 1]
   Category: ☐ Market ☐ Technology ☐ Regulatory ☐ Social ☐ Competitive ☐ Economic
   Description: [Detailed explanation of the opportunity]
   Market Size/Potential: [Revenue potential or impact]
   Timeframe: [When we could/should capitalize]
   Alignment with Strengths: [Which strengths help us capture this]
   Investment Required: [Resources needed]
   Feasibility: ☐ High ☐ Medium ☐ Low
   Priority: ☐ High ☐ Medium ☐ Low

2. [Opportunity 2]
   Category: ☐ Market ☐ Technology ☐ Regulatory ☐ Social ☐ Competitive ☐ Economic
   Description: [Detailed explanation]
   Market Size/Potential: [Value]
   Timeframe: [Timeline]
   Alignment with Strengths: [Connection]
   Investment Required: [Cost]
   Feasibility: ☐ High ☐ Medium ☐ Low
   Priority: ☐ High ☐ Medium ☐ Low

3. [Opportunity 3]
   Category: ☐ Market ☐ Technology ☐ Regulatory ☐ Social ☐ Competitive ☐ Economic
   Description: [Detailed explanation]
   Market Size/Potential: [Value]
   Timeframe: [Timeline]
   Alignment with Strengths: [Connection]
   Investment Required: [Cost]
   Feasibility: ☐ High ☐ Medium ☐ Low
   Priority: ☐ High ☐ Medium ☐ Low

[Continue for 5-10 total opportunities]

OPPORTUNITIES BY CATEGORY:

Market Trends:
• [Emerging trend we can leverage]
• [Market shift in our favor]

Customer Needs:
• [Unmet need we can address]
• [Growing demand we can meet]

Technology Advances:
• [New technology to adopt]
• [Tech-enabled opportunity]

Partnerships & Alliances:
• [Potential strategic partnership]
• [Co-marketing opportunity]

Geographic Expansion:
• [New market to enter]
• [Region to expand into]

Product/Service Innovation:
• [New product opportunity]
• [Service expansion possibility]

Regulatory/Policy Changes:
• [Favorable regulation]
• [Policy change creating opportunity]

Competitor Gaps:
• [Area where competitors are weak]
• [Underserved segment]

Economic Factors:
• [Economic trend favoring us]
• [Consumer spending pattern]

---

THREATS (What external challenges could harm us?)

1. [Threat 1]
   Category: ☐ Competition ☐ Market ☐ Technology ☐ Regulatory ☐ Economic ☐ Social
   Description: [Detailed explanation of the threat]
   Likelihood: ☐ High ☐ Medium ☐ Low
   Potential Impact: ☐ Severe ☐ Moderate ☐ Minor
   Timeframe: [When this could materialize]
   Early Warning Signs: [Indicators to monitor]
   Mitigation Strategy: [How we can prepare/respond]
   Priority Level: ☐ Critical ☐ High ☐ Medium ☐ Low

2. [Threat 2]
   Category: ☐ Competition ☐ Market ☐ Technology ☐ Regulatory ☐ Economic ☐ Social
   Description: [Detailed explanation]
   Likelihood: ☐ High ☐ Medium ☐ Low
   Potential Impact: ☐ Severe ☐ Moderate ☐ Minor
   Timeframe: [Timeline]
   Early Warning Signs: [Indicators]
   Mitigation Strategy: [Response plan]
   Priority Level: ☐ Critical ☐ High ☐ Medium ☐ Low

3. [Threat 3]
   Category: ☐ Competition ☐ Market ☐ Technology ☐ Regulatory ☐ Economic ☐ Social
   Description: [Detailed explanation]
   Likelihood: ☐ High ☐ Medium ☐ Low
   Potential Impact: ☐ Severe ☐ Moderate ☐ Minor
   Timeframe: [Timeline]
   Early Warning Signs: [Indicators]
   Mitigation Strategy: [Response plan]
   Priority Level: ☐ Critical ☐ High ☐ Medium ☐ Low

[Continue for 5-10 total threats]

THREATS BY CATEGORY:

Competitive Threats:
• [New competitor entering]
• [Competitor innovation]
• [Price war risk]

Market Changes:
• [Shifting customer preferences]
• [Market saturation]
• [Declining demand]

Technological Disruption:
• [Disruptive technology]
• [Platform shift]
• [Automation threat]

Regulatory Risks:
• [New regulation]
• [Compliance requirement]
• [Policy change]

Economic Conditions:
• [Recession risk]
• [Inflation impact]
• [Currency fluctuation]

Supply Chain:
• [Supplier dependency]
• [Material shortage]
• [Logistics disruption]

Talent & Resources:
• [Skill shortage]
• [Talent competition]
• [Key person risk]

Reputation:
• [PR risk]
• [Social media backlash]
• [Trust erosion]

---

SWOT MATRIX - STRATEGIC ANALYSIS:

S-O STRATEGIES (Use Strengths to Capitalize on Opportunities):

1. [Strategy combining specific strength with specific opportunity]
   Strength Used: [Which strength]
   Opportunity Captured: [Which opportunity]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

2. [Strategy combining strength with opportunity]
   Strength Used: [Which strength]
   Opportunity Captured: [Which opportunity]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

3. [Strategy combining strength with opportunity]
   Strength Used: [Which strength]
   Opportunity Captured: [Which opportunity]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

W-O STRATEGIES (Overcome Weaknesses to Pursue Opportunities):

1. [Strategy to fix weakness and capture opportunity]
   Weakness to Address: [Which weakness]
   Opportunity to Pursue: [Which opportunity]
   Improvement Needed: [What must change]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

2. [Strategy to fix weakness and capture opportunity]
   Weakness to Address: [Which weakness]
   Opportunity to Pursue: [Which opportunity]
   Improvement Needed: [What must change]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

S-T STRATEGIES (Use Strengths to Minimize Threats):

1. [Strategy using strength to counter threat]
   Strength to Leverage: [Which strength]
   Threat to Counter: [Which threat]
   Defensive Action: [How strength protects us]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

2. [Strategy using strength to counter threat]
   Strength to Leverage: [Which strength]
   Threat to Counter: [Which threat]
   Defensive Action: [How strength protects us]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

W-T STRATEGIES (Minimize Weaknesses and Avoid Threats):

1. [Defensive strategy addressing weakness and threat]
   Weakness to Minimize: [Which weakness]
   Threat to Avoid: [Which threat]
   Protective Action: [How to defend]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

2. [Defensive strategy addressing weakness and threat]
   Weakness to Minimize: [Which weakness]
   Threat to Avoid: [Which threat]
   Protective Action: [How to defend]
   Expected Outcome: [Result]
   Action Steps: [How to execute]

---

PRIORITY ACTIONS (Top 3 in each category):

TOP 3 OPPORTUNITIES TO PURSUE:

1. [Opportunity Name]
   Why: [Strategic rationale]
   Action Plan: [Steps to capture it]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure success]

2. [Opportunity Name]
   Why: [Strategic rationale]
   Action Plan: [Steps to capture it]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure success]

3. [Opportunity Name]
   Why: [Strategic rationale]
   Action Plan: [Steps to capture it]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure success]

TOP 3 WEAKNESSES TO ADDRESS:

1. [Weakness to Fix]
   Why: [Impact if not fixed]
   Improvement Plan: [Steps to address it]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure improvement]

2. [Weakness to Fix]
   Why: [Impact if not fixed]
   Improvement Plan: [Steps to address it]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure improvement]

3. [Weakness to Fix]
   Why: [Impact if not fixed]
   Improvement Plan: [Steps to address it]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure improvement]

TOP 3 THREATS TO MITIGATE:

1. [Threat to Address]
   Why: [Potential impact]
   Mitigation Plan: [Steps to reduce risk]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure mitigation]

2. [Threat to Address]
   Why: [Potential impact]
   Mitigation Plan: [Steps to reduce risk]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure mitigation]

3. [Threat to Address]
   Why: [Potential impact]
   Mitigation Plan: [Steps to reduce risk]
   Owner: [Responsible person/team]
   Timeline: [Deadline]
   Resources Needed: [Budget, people, tools]
   Success Metric: [How we'll measure mitigation]

---

STRATEGIC RECOMMENDATIONS:

Based on this SWOT analysis, our strategic priorities are:

SHORT-TERM (0-6 months):
1. [Immediate action/focus]
2. [Immediate action/focus]
3. [Immediate action/focus]

MEDIUM-TERM (6-18 months):
1. [Medium-term initiative]
2. [Medium-term initiative]
3. [Medium-term initiative]

LONG-TERM (18+ months):
1. [Long-term strategic direction]
2. [Long-term strategic direction]
3. [Long-term strategic direction]

---

COMPETITIVE POSITIONING:

Current Market Position:
[Description of where we stand relative to competitors]

Recommended Strategic Positioning:

☐ Cost Leadership (Compete on price/efficiency)
☐ Differentiation (Compete on unique value/quality)
☐ Focus/Niche Strategy (Dominate specific segment)
☐ Blue Ocean (Create new market category)
☐ Fast Follower (Quickly adapt innovation)

Rationale for Chosen Position:
[Why this positioning makes sense given our SWOT]

---

SUCCESS METRICS & MONITORING:

How we'll track progress on SWOT-based initiatives:

Quarterly Metrics:
• [Metric 1]: Current [X], Target [Y]
• [Metric 2]: Current [X], Target [Y]
• [Metric 3]: Current [X], Target [Y]
• [Metric 4]: Current [X], Target [Y]

Annual Metrics:
• [Metric 1]: Current [X], Target [Y]
• [Metric 2]: Current [X], Target [Y]
• [Metric 3]: Current [X], Target [Y]

Review Schedule:
• Monthly: Progress check on priority actions
• Quarterly: Update SWOT based on changes
• Annually: Comprehensive SWOT refresh

Next SWOT Review Date: [Date]

---

KEY INSIGHTS & TAKEAWAYS:

Most Critical Findings:
• [Key insight 1]
• [Key insight 2]
• [Key insight 3]

Biggest Opportunities:
• [Opportunity with highest potential]

Biggest Threats:
• [Threat requiring most attention]

Critical Gaps to Address:
• [Gap between current and desired state]

---

ASSUMPTIONS & CAVEATS:

Key Assumptions Made:
• [Assumption 1 that underlies this analysis]
• [Assumption 2 that underlies this analysis]
• [Assumption 3 that underlies this analysis]

Limitations of This Analysis:
• [Limitation 1]
• [Limitation 2]

Additional Research Needed:
• [Research topic 1]
• [Research topic 2]
• [Research topic 3]

---

APPROVALS & DISTRIBUTION:

Prepared by: _________________ Date: _______
[Analyst Name, Title]

Reviewed by: _________________ Date: _______
[Manager Name, Title]

Approved by: _________________ Date: _______
[Executive Name, Title]

Distribution List:
☐ Executive Team
☐ Marketing Team
☐ Sales Team
☐ Product Team
☐ Finance Team
☐ Board of Directors
☐ Other: [Specify]

NOTES:
[Additional context, observations, or considerations]
    `
  },
  {
    id: 'marketing-customer-persona',
    title: 'Customer Persona Worksheet',
    description: 'Detailed customer persona template for understanding and targeting your ideal customer',
    category: 'marketing',
    downloadCount: 4830,
    fileSize: '32.6 KB',
    rating: 4.7,
    tags: ['persona', 'customer', 'targeting', 'audience', 'research'],
    createdAt: '2024-01-19T08:30:00Z',
    updatedAt: '2024-02-05T16:00:00Z',
    content: `
CUSTOMER PERSONA WORKSHEET

PERSONA OVERVIEW:
Persona Name: [Give them a memorable name, e.g., "Marketing Mary" or "Developer Dan"]
Date Created: [Date]
Created By: [Your name]
Last Updated: [Date]
Version: [1.0]

This persona represents: [Brief description of who this persona represents]
Estimated % of customer base: [X%]

DEMOGRAPHIC INFORMATION:

Age: [Range, e.g., 28-35]
Gender: [If relevant to your business]
Location: [City type, region, country]
Education: [Highest level of education]
Occupation: [Job title/role]
Industry: [Industry they work in]
Company Size: [If B2B - Small/Medium/Large/Enterprise]
Income Level: [Range]
Marital Status: [If relevant]
Family Status: [e.g., "Has 2 kids" if relevant]

PROFESSIONAL BACKGROUND:

Job Title: [Specific title]
Department: [Department]
Career Level: [Entry/Mid/Senior/C-Level]
Years of Experience: [Range]
Reporting Structure: [Who they report to, who reports to them]
Team Size: [If they manage a team]

Key Responsibilities:
• [Responsibility 1]
• [Responsibility 2]
• [Responsibility 3]
• [Responsibility 4]

Daily Tasks:
• [Task 1]
• [Task 2]
• [Task 3]

Skills & Expertise:
• [Skill 1]
• [Skill 2]
• [Skill 3]

Tools They Use:
• [Tool/Software 1]
• [Tool/Software 2]
• [Tool/Software 3]

PSYCHOGRAPHIC PROFILE:

Personality Traits:
• [Trait 1 - e.g., "Detail-oriented"]
• [Trait 2 - e.g., "Risk-averse"]
• [Trait 3 - e.g., "Collaborative"]

Values & Beliefs:
• [Value 1]
• [Value 2]
• [Value 3]

Interests & Hobbies:
• [Interest 1]
• [Interest 2]
• [Interest 3]

Lifestyle:
• [Lifestyle characteristic 1]
• [Lifestyle characteristic 2]

Motivations (What drives them):
• [Motivation 1]
• [Motivation 2]
• [Motivation 3]

Fears & Frustrations:
• [Fear/Frustration 1]
• [Fear/Frustration 2]
• [Fear/Frustration 3]

GOALS & CHALLENGES:

Primary Goals:
Personal Goals:
• [Goal 1]
• [Goal 2]
• [Goal 3]

Professional Goals:
• [Goal 1]
• [Goal 2]
• [Goal 3]

Main Challenges & Pain Points:
Challenge 1: [Challenge name]
Description: [Detailed description]
Impact: [How this affects them]
Current Solution: [What they do now to address it]
Satisfaction with Current Solution: [Low/Medium/High]

Challenge 2: [Challenge name]
Description: [Detailed description]
Impact: [How this affects them]
Current Solution: [What they do now to address it]
Satisfaction with Current Solution: [Low/Medium/High]

Challenge 3: [Challenge name]
Description: [Detailed description]
Impact: [How this affects them]
Current Solution: [What they do now to address it]
Satisfaction with Current Solution: [Low/Medium/High]

Barriers to Success:
• [Barrier 1]
• [Barrier 2]
• [Barrier 3]

INFORMATION & MEDIA CONSUMPTION:

Preferred Communication Channels:
• [Channel 1 - e.g., Email]
• [Channel 2 - e.g., LinkedIn]
• [Channel 3 - e.g., Phone]

Communication Style:
☐ Formal
☐ Casual
☐ Direct and brief
☐ Detailed and comprehensive
☐ Visual (prefers images/videos)
☐ Text-based (prefers written content)

Information Sources:
Where they learn about new products/services:
• [Source 1 - e.g., Industry blogs]
• [Source 2 - e.g., Peer recommendations]
• [Source 3 - e.g., LinkedIn]
• [Source 4 - e.g., Trade shows]

Trusted Influencers:
• [Influencer/Thought leader 1]
• [Influencer/Thought leader 2]
• [Influencer/Thought leader 3]

Publications They Read:
• [Publication 1]
• [Publication 2]
• [Publication 3]

Podcasts They Listen To:
• [Podcast 1]
• [Podcast 2]

Social Media Platforms (Ranked by usage):
1. [Platform 1] - [How they use it]
2. [Platform 2] - [How they use it]
3. [Platform 3] - [How they use it]

Online Communities/Forums:
• [Community 1]
• [Community 2]

Events They Attend:
• [Conference/Event 1]
• [Conference/Event 2]

Content Preferences:
☐ Blog posts
☐ Videos
☐ Podcasts
☐ Infographics
☐ Webinars
☐ Case studies
☐ Whitepapers/eBooks
☐ Industry reports
☐ Email newsletters

Preferred Content Length:
☐ Quick tips (1-2 min)
☐ Medium-depth (5-10 min)
☐ Comprehensive guides (20+ min)

BUYING BEHAVIOR:

Decision-Making Process:
Decision-Making Style:
☐ Analytical (data-driven)
☐ Intuitive (gut feeling)
☐ Collaborative (seeks input)
☐ Quick (decisive)
☐ Deliberate (takes time)

Who Else is Involved in Purchase Decisions:
• [Role 1 - e.g., "IT Manager"]
• [Role 2 - e.g., "CFO"]
• [Role 3 - e.g., "End users"]

Their Role in Purchase:
☐ End user
☐ Influencer
☐ Decision maker
☐ Budget holder
☐ Gatekeeper

Typical Buying Process:
1. [Step 1 - e.g., "Identifies problem"]
2. [Step 2 - e.g., "Researches solutions online"]
3. [Step 3 - e.g., "Compares 3-4 vendors"]
4. [Step 4 - e.g., "Requests demos"]
5. [Step 5 - e.g., "Gets approval from manager"]
6. [Step 6 - e.g., "Makes purchase"]

Average Decision Timeline:
[X days/weeks/months from awareness to purchase]

Research Behavior:
Before making a purchase, they typically:
• [Behavior 1 - e.g., "Read 3-5 reviews"]
• [Behavior 2 - e.g., "Compare pricing"]
• [Behavior 3 - e.g., "Request ROI calculation"]

Key Decision Criteria (Ranked):
1. [Criterion 1 - e.g., "Price"]
2. [Criterion 2 - e.g., "Ease of use"]
3. [Criterion 3 - e.g., "Integration capabilities"]
4. [Criterion 4 - e.g., "Customer support"]
5. [Criterion 5 - e.g., "Brand reputation"]

Budget Constraints:
Typical Budget Range: $[Low] - $[High]
Budget Approval Required: [Yes/No]
Budget Cycle: [Annual/Quarterly/As needed]

OBJECTIONS & CONCERNS:

Common Objections to Our Product/Service:
Objection 1: [Objection]
Why they feel this way: [Reason]
How to address: [Counter-argument]

Objection 2: [Objection]
Why they feel this way: [Reason]
How to address: [Counter-argument]

Objection 3: [Objection]
Why they feel this way: [Reason]
How to address: [Counter-argument]

Concerns About:
• Price: [Specific concern]
• Implementation: [Specific concern]
• Learning Curve: [Specific concern]
• ROI: [Specific concern]
• Support: [Specific concern]

Deal Breakers:
• [Deal breaker 1]
• [Deal breaker 2]
• [Deal breaker 3]

RELATIONSHIP WITH OUR BRAND:

Current Relationship:
☐ Unaware of us
☐ Aware but never engaged
☐ Has visited our website
☐ Downloaded content from us
☐ Trial user
☐ Current customer
☐ Former customer
☐ Detractor

How They Found Us:
[Primary discovery method]

Brand Perception:
What they think about our brand: [Perception]

Competitors They Consider:
1. [Competitor 1] - Why they like them: [Reason]
2. [Competitor 2] - Why they like them: [Reason]
3. [Competitor 3] - Why they like them: [Reason]

Why They Would Choose Us Over Competitors:
• [Reason 1]
• [Reason 2]
• [Reason 3]

Why They Might Choose Competitors Instead:
• [Reason 1]
• [Reason 2]

PRODUCT/SERVICE USAGE:

How They Would Use Our Product/Service:
Primary Use Case: [Main way they'd use it]
Secondary Use Cases:
• [Use case 1]
• [Use case 2]

Expected Outcomes:
• [Outcome 1]
• [Outcome 2]
• [Outcome 3]

Features Most Important to Them:
1. [Feature 1]
2. [Feature 2]
3. [Feature 3]
4. [Feature 4]
5. [Feature 5]

Features They Don't Care About:
• [Feature 1]
• [Feature 2]

Success Metrics They Care About:
• [Metric 1]
• [Metric 2]
• [Metric 3]

IDEAL CUSTOMER EXPERIENCE:

Preferred Sales Process:
☐ Self-serve (no sales contact)
☐ Brief consultation
☐ In-depth consultative selling
☐ Product demo required
☐ Free trial first

Onboarding Preferences:
☐ Self-guided tutorial
☐ Video training
☐ Live training session
☐ Dedicated account manager
☐ Just-in-time help

Support Expectations:
Preferred support channels:
• [Channel 1]
• [Channel 2]

Expected response time: [Timeframe]
Preferred support hours: [Hours]

What Makes Them a Loyal Customer:
• [Factor 1]
• [Factor 2]
• [Factor 3]

What Would Make Them Leave:
• [Factor 1]
• [Factor 2]
• [Factor 3]

MARKETING MESSAGING FOR THIS PERSONA:

Value Proposition (What resonates):
[Tailored value proposition that speaks to their specific needs]

Key Messages:
1. [Message 1 addressing their top pain point]
2. [Message 2 addressing their goal]
3. [Message 3 building trust/credibility]

Emotional Triggers:
• [Trigger 1 - e.g., "Fear of falling behind"]
• [Trigger 2 - e.g., "Desire for recognition"]
• [Trigger 3 - e.g., "Need for security"]

Words & Phrases That Resonate:
• [Word/phrase 1]
• [Word/phrase 2]
• [Word/phrase 3]

Words & Phrases to Avoid:
• [Word/phrase 1 - why to avoid]
• [Word/phrase 2 - why to avoid]

Content Topics That Interest Them:
• [Topic 1]
• [Topic 2]
• [Topic 3]
• [Topic 4]

REAL PERSONA EXAMPLE:

Meet [First Name Last Name]:
[2-3 paragraph narrative story about a real or composite person that embodies this persona]

A Day in Their Life:
Morning: [What they do]
Midday: [What they do]
Afternoon: [What they do]
Evening: [What they do]

Quote (In their words):
"[A quote that captures their main frustration or goal]"

PERSONA VALIDATION:

Data Sources Used:
☐ Customer interviews ([#] conducted)
☐ Surveys ([#] responses)
☐ Analytics data
☐ Sales team input
☐ Customer support tickets
☐ Social media listening
☐ Market research
☐ Competitor analysis

Confidence Level:
☐ High (validated with substantial data)
☐ Medium (based on good data but needs validation)
☐ Low (mostly assumptions, needs research)

Next Steps for Validation:
• [Action 1]
• [Action 2]
• [Action 3]

HOW TO USE THIS PERSONA:

Marketing Applications:
• Content creation: [How to tailor content]
• Ad targeting: [How to target ads]
• Email campaigns: [How to segment]
• Social media: [Which platforms and messaging]

Sales Applications:
• Lead qualification: [What makes them qualified]
• Sales pitch: [Key points to emphasize]
• Demo focus: [Features to highlight]
• Objection handling: [Common objections and responses]

Product Development:
• Feature prioritization: [What features matter most]
• UX considerations: [How they prefer to interact]
• Documentation needs: [Support content needed]

Customer Success:
• Onboarding focus: [What to emphasize]
• Engagement strategy: [How to keep them engaged]
• Expansion opportunities: [Upsell/cross-sell potential]

PERSONA SNAPSHOT (One-Pager Summary):

[Create a visual-friendly one-page summary with the most critical information]

WHO THEY ARE:
• [Key demographic]
• [Key professional info]

GOALS:
• [Top goal 1]
• [Top goal 2]

CHALLENGES:
• [Top pain point 1]
• [Top pain point 2]

HOW WE HELP:
[One sentence value proposition]

WHERE TO REACH THEM:
• [Channel 1]
• [Channel 2]

WHAT THEY CARE ABOUT WHEN BUYING:
• [Key criterion 1]
• [Key criterion 2]

---

NOTES & INSIGHTS:
[Additional observations, insights, or context about this persona]

PERSONA OWNER: _________________
[Name of person responsible for keeping this persona updated]

NEXT REVIEW DATE: _______
[When to review and update this persona - recommend quarterly]
    `
  },
  {
    id: 'marketing-budget-template',
    title: 'Marketing Budget Template',
    description: 'Comprehensive marketing budget template with categories, tracking, and ROI analysis',
    category: 'marketing',
    downloadCount: 5210,
    fileSize: '35.7 KB',
    rating: 4.8,
    tags: ['budget', 'finance', 'planning', 'roi', 'spending'],
    createdAt: '2024-01-21T10:00:00Z',
    updatedAt: '2024-02-06T14:30:00Z',
    content: `
MARKETING BUDGET TEMPLATE

BUDGET OVERVIEW:
Company: [Company Name]
Fiscal Year/Period: [FY 2024, Q1 2024, etc.]
Budget Owner: [Marketing Manager Name]
Date Created: [Date]
Last Updated: [Date]
Status: [Draft / Submitted / Approved / In Progress]

EXECUTIVE SUMMARY:
Total Marketing Budget: $[Amount]
% of Company Revenue: [X%]
Year-over-Year Change: [+/- X%]

Primary Goals:
• [Goal 1 with target metric]
• [Goal 2 with target metric]
• [Goal 3 with target metric]

---

BUDGET BREAKDOWN BY CATEGORY:

CATEGORY 1: DIGITAL ADVERTISING ($[Amount] | [%]% of total budget)

Paid Search (Google Ads, Bing Ads):
Monthly Budget: $[Amount]
Annual Budget: $[Amount x 12]
Expected Leads: [#]
Expected Cost Per Lead: $[Amount]
Expected ROI: [X:1]

Paid Social Media:
Facebook/Instagram Ads: $[Amount]/month = $[Annual]
LinkedIn Ads: $[Amount]/month = $[Annual]
Twitter Ads: $[Amount]/month = $[Annual]
TikTok Ads: $[Amount]/month = $[Annual]
Pinterest Ads: $[Amount]/month = $[Annual]
Total Paid Social: $[Annual total]

Display Advertising:
Programmatic Display: $[Amount]/month = $[Annual]
Banner Ads: $[Amount]/month = $[Annual]
Native Advertising: $[Amount]/month = $[Annual]
Total Display: $[Annual total]

Retargeting/Remarketing:
Budget: $[Amount]/month = $[Annual]
Expected Conversion Rate: [%]

Video Advertising:
YouTube Ads: $[Amount]/month = $[Annual]
Connected TV/OTT: $[Amount]/month = $[Annual]
Total Video: $[Annual total]

Digital Advertising Total: $[Amount]

---

CATEGORY 2: CONTENT MARKETING ($[Amount] | [%]% of total budget)

Content Creation:
Blog Posts: $[Amount]/month = $[Annual]
  ([X] posts per month @ $[cost] per post)
Ebooks/Guides: $[Amount per quarter] = $[Annual]
Whitepapers: $[Amount per quarter] = $[Annual]
Case Studies: $[Amount each] x [#] = $[Annual]
Infographics: $[Amount each] x [#] = $[Annual]
Videos: $[Amount each] x [#] = $[Annual]
Podcasts: $[Amount]/month = $[Annual]

Content Distribution:
Sponsored content: $[Amount]/month = $[Annual]
Content syndication: $[Amount]/month = $[Annual]

Freelance Writers/Creators:
Budget: $[Amount]/month = $[Annual]

Photography/Videography:
Stock photos/videos: $[Amount]/month = $[Annual]
Custom shoots: $[Amount per quarter] = $[Annual]

Content Marketing Total: $[Amount]

---

CATEGORY 3: SOCIAL MEDIA MARKETING ($[Amount] | [%]% of total budget)

Organic Social Media:
Social media management tools: $[Amount]/month = $[Annual]
  (Hootsuite, Buffer, Sprout Social, etc.)
Content creation for social: $[Amount]/month = $[Annual]
Community management: $[Amount]/month = $[Annual]

Paid Social (See Digital Advertising section): $[Amount]

Influencer Marketing:
Micro-influencers: $[Amount per campaign] x [#] = $[Annual]
Macro-influencers: $[Amount per campaign] x [#] = $[Annual]
Brand ambassadors: $[Amount]/month = $[Annual]

Social Media Total: $[Amount]

---

CATEGORY 4: EMAIL MARKETING ($[Amount] | [%]% of total budget)

Email Service Provider:
Platform (Mailchimp, HubSpot, etc.): $[Amount]/month = $[Annual]

Email Design & Development:
Templates: $[Amount]/month = $[Annual]
Custom designs: $[Amount per quarter] = $[Annual]

List Building:
Lead magnets creation: $[Amount]/month = $[Annual]
List acquisition: $[Amount]/month = $[Annual]

Email Marketing Total: $[Amount]

---

CATEGORY 5: SEO & WEBSITE ($[Amount] | [%]% of total budget)

SEO Services:
SEO tools (Ahrefs, SEMrush, etc.): $[Amount]/month = $[Annual]
SEO agency/consultant: $[Amount]/month = $[Annual]
Content optimization: $[Amount]/month = $[Annual]
Link building: $[Amount]/month = $[Annual]

Website:
Web hosting: $[Amount]/month = $[Annual]
Domain registration/renewal: $[Amount]/year
Website maintenance: $[Amount]/month = $[Annual]
Website updates/improvements: $[Amount per quarter] = $[Annual]
Landing page builder (Unbounce, etc.): $[Amount]/month = $[Annual]
CMS (WordPress, etc.): $[Amount]/month = $[Annual]
A/B testing tools (Optimizely, VWO): $[Amount]/month = $[Annual]

SEO & Website Total: $[Amount]

---

CATEGORY 6: MARKETING TECHNOLOGY & TOOLS ($[Amount] | [%]% of total budget)

CRM Platform:
Software (HubSpot, Salesforce, etc.): $[Amount]/month = $[Annual]

Marketing Automation:
Platform: $[Amount]/month = $[Annual]

Analytics & Reporting:
Google Analytics 360 (if paid): $[Amount]/year
Other analytics tools: $[Amount]/month = $[Annual]
Business intelligence tools: $[Amount]/month = $[Annual]

Design & Creative Tools:
Adobe Creative Cloud: $[Amount]/month = $[Annual]
Canva Pro: $[Amount]/month = $[Annual]
Stock image subscriptions: $[Amount]/month = $[Annual]

Project Management:
Tools (Asana, Monday.com, etc.): $[Amount]/month = $[Annual]

Other Software:
[Tool name]: $[Amount]/month = $[Annual]
[Tool name]: $[Amount]/month = $[Annual]

Marketing Technology Total: $[Amount]

---

CATEGORY 7: EVENTS & TRADE SHOWS ($[Amount] | [%]% of total budget)

Trade Shows & Conferences:
[Event 1]: $[Booth + travel + materials] = $[Amount]
[Event 2]: $[Booth + travel + materials] = $[Amount]
[Event 3]: $[Booth + travel + materials] = $[Amount]

Booth Design & Materials:
Booth design/build: $[Amount]
Banner stands & displays: $[Amount]
Promotional materials: $[Amount]
Giveaways: $[Amount]

Sponsored Events:
[Event sponsorship 1]: $[Amount]
[Event sponsorship 2]: $[Amount]

Virtual Events:
Webinar platform: $[Amount]/month = $[Annual]
Virtual conference sponsorships: $[Amount]

Hosted Events:
Customer appreciation events: $[Amount per quarter] = $[Annual]
Networking events: $[Amount per quarter] = $[Annual]

Events Total: $[Amount]

---

CATEGORY 8: PUBLIC RELATIONS & MEDIA ($[Amount] | [%]% of total budget)

PR Agency/Consultant:
Retainer: $[Amount]/month = $[Annual]

Press Release Distribution:
Service (PR Newswire, etc.): $[Amount per release] x [#] = $[Annual]

Media Monitoring:
Monitoring service: $[Amount]/month = $[Annual]

Crisis Communication:
Reserved budget: $[Amount]

PR & Media Total: $[Amount]

---

CATEGORY 9: TRADITIONAL ADVERTISING ($[Amount] | [%]% of total budget)

Print Advertising:
Magazine ads: $[Amount per quarter] = $[Annual]
Newspaper ads: $[Amount per quarter] = $[Annual]
Industry publications: $[Amount]/month = $[Annual]

Broadcast:
Radio spots: $[Amount per quarter] = $[Annual]
TV commercials: $[Amount per quarter] = $[Annual]

Outdoor:
Billboards: $[Amount]/month = $[Annual]
Transit advertising: $[Amount per quarter] = $[Annual]

Direct Mail:
Production: $[Amount per campaign] x [#] = $[Annual]
Postage: $[Amount per campaign] x [#] = $[Annual]
List rental: $[Amount per campaign] x [#] = $[Annual]

Traditional Advertising Total: $[Amount]

---

CATEGORY 10: PARTNERSHIPS & SPONSORSHIPS ($[Amount] | [%]% of total budget)

Strategic Partnerships:
Co-marketing agreements: $[Amount]
Affiliate program commissions: $[Amount]/month = $[Annual]

Sponsorships:
Podcast sponsorships: $[Amount each] x [#] = $[Annual]
Newsletter sponsorships: $[Amount each] x [#] = $[Annual]
Community sponsorships: $[Amount]

Partnerships Total: $[Amount]

---

CATEGORY 11: BRAND & CREATIVE ($[Amount] | [%]% of total budget)

Brand Development:
Brand refresh/update: $[Amount]
Brand guidelines: $[Amount]

Creative Services:
Graphic design: $[Amount]/month = $[Annual]
Copywriting: $[Amount]/month = $[Annual]
Creative agency retainer: $[Amount]/month = $[Annual]

Brand & Creative Total: $[Amount]

---

CATEGORY 12: MARKET RESEARCH ($[Amount] | [%]% of total budget)

Research Services:
Customer surveys: $[Amount per quarter] = $[Annual]
Focus groups: $[Amount each] x [#] = $[Annual]
Market research reports: $[Amount]
Competitive intelligence: $[Amount]/month = $[Annual]

Survey Tools:
Software (SurveyMonkey, Qualtrics): $[Amount]/month = $[Annual]

Market Research Total: $[Amount]

---

CATEGORY 13: STAFFING & PROFESSIONAL SERVICES ($[Amount] | [%]% of total budget)

Freelancers & Contractors:
Content writers: $[Amount]/month = $[Annual]
Designers: $[Amount]/month = $[Annual]
Developers: $[Amount]/month = $[Annual]
Social media managers: $[Amount]/month = $[Annual]

Agencies:
Digital marketing agency: $[Amount]/month = $[Annual]
SEO agency: $[Amount]/month = $[Annual]
PR agency: (see PR section) $[Amount]

Consultants:
Marketing strategy consultant: $[Amount]
Specialized consultants: $[Amount]

Training & Development:
Conferences & workshops: $[Amount per person] x [#] = $[Amount]
Online courses: $[Amount]
Certifications: $[Amount]

Staffing Total: $[Amount]

---

CATEGORY 14: MISCELLANEOUS & CONTINGENCY ($[Amount] | [%]% of total budget)

Miscellaneous Expenses:
Subscriptions not listed elsewhere: $[Amount]/month = $[Annual]
Testing new channels: $[Amount]
Unexpected opportunities: $[Amount]

Contingency Fund (10% recommended):
$[Amount]

Miscellaneous Total: $[Amount]

---

TOTAL ANNUAL MARKETING BUDGET: $[GRAND TOTAL]

---

BUDGET ALLOCATION SUMMARY:

Category | Budget | % of Total | Monthly Average
---------|--------|------------|------------------
Digital Advertising | $[Amount] | [%]% | $[Amount]
Content Marketing | $[Amount] | [%]% | $[Amount]
Social Media | $[Amount] | [%]% | $[Amount]
Email Marketing | $[Amount] | [%]% | $[Amount]
SEO & Website | $[Amount] | [%]% | $[Amount]
Marketing Technology | $[Amount] | [%]% | $[Amount]
Events & Trade Shows | $[Amount] | [%]% | $[Amount]
PR & Media | $[Amount] | [%]% | $[Amount]
Traditional Advertising | $[Amount] | [%]% | $[Amount]
Partnerships | $[Amount] | [%]% | $[Amount]
Brand & Creative | $[Amount] | [%]% | $[Amount]
Market Research | $[Amount] | [%]% | $[Amount]
Staffing | $[Amount] | [%]% | $[Amount]
Miscellaneous | $[Amount] | [%]% | $[Amount]
---------|--------|------------|------------------
TOTAL | $[Amount] | 100% | $[Amount]

---

MONTHLY BUDGET BREAKDOWN:

Month | Planned Spend | Actual Spend | Variance | Notes
------|---------------|--------------|----------|-------
January | $[Amount] | $[Amount] | $[Difference] | [Notes]
February | $[Amount] | $[Amount] | $[Difference] | [Notes]
March | $[Amount] | $[Amount] | $[Difference] | [Notes]
April | $[Amount] | $[Amount] | $[Difference] | [Notes]
May | $[Amount] | $[Amount] | $[Difference] | [Notes]
June | $[Amount] | $[Amount] | $[Difference] | [Notes]
July | $[Amount] | $[Amount] | $[Difference] | [Notes]
August | $[Amount] | $[Amount] | $[Difference] | [Notes]
September | $[Amount] | $[Amount] | $[Difference] | [Notes]
October | $[Amount] | $[Amount] | $[Difference] | [Notes]
November | $[Amount] | $[Amount] | $[Difference] | [Notes]
December | $[Amount] | $[Amount] | $[Difference] | [Notes]
---------|--------|--------------|----------
TOTAL | $[Amount] | $[Amount] | $[Difference] |

---

EXPECTED ROI & PERFORMANCE METRICS:

Expected Marketing ROI: [X:1]
Expected Customer Acquisition Cost (CAC): $[Amount]
Expected Customer Lifetime Value (CLV): $[Amount]
Expected LTV:CAC Ratio: [X:1]

Expected Results:
• New Leads: [#]
• Marketing Qualified Leads (MQLs): [#]
• Sales Qualified Leads (SQLs): [#]
• New Customers: [#]
• Revenue Attributed to Marketing: $[Amount]
• Website Visitors: [#]/month
• Email List Growth: [#] new subscribers
• Social Media Followers: [#] total
• Brand Awareness Lift: [X%]

---

BUDGET JUSTIFICATION:

Why This Budget is Necessary:
[Explanation of business goals and how this budget supports them]

Expected Impact on Business:
• [Impact 1]
• [Impact 2]
• [Impact 3]

Comparison to Industry Benchmarks:
Industry Average Marketing Budget: [X% of revenue]
Our Budget: [X% of revenue]
Variance: [Higher/Lower than average because...]

---

RISK ASSESSMENT:

Budget Risks:
Risk 1: [Potential issue]
Likelihood: [High/Medium/Low]
Mitigation: [How to address]

Risk 2: [Potential issue]
Likelihood: [High/Medium/Low]
Mitigation: [How to address]

Opportunities for Reallocation:
• [Channel or category that could be reduced]
• [Channel or category that should get more investment]

---

APPROVALS:

Budget Prepared by: _________________ Date: _______
[Marketing Manager]

Reviewed by: _________________ Date: _______
[Marketing Director/VP]

Approved by CFO: _________________ Date: _______
[CFO Name]

Final Approval by CEO: _________________ Date: _______
[CEO Name]

---

TRACKING & REPORTING:

Budget Review Frequency:
☐ Weekly
☐ Bi-weekly
☐ Monthly
☐ Quarterly

Reporting Dashboard:
[Tool/platform used for tracking]

Key Reports:
• Monthly budget vs. actual report
• ROI by channel report
• CAC trend report
• Campaign performance report

Budget Owner: _________________
[Person responsible for managing and tracking the budget]

---

NOTES:
[Additional context, assumptions, or important information about this budget]
    `
  },
  {
    id: 'marketing-creative-brief',
    title: 'Creative Brief',
    description: 'Comprehensive creative brief template for marketing campaigns, ads, and design projects',
    category: 'marketing',
    downloadCount: 5580,
    fileSize: '29.4 KB',
    rating: 4.8,
    tags: ['creative', 'brief', 'campaign', 'design', 'advertising'],
    createdAt: '2024-01-23T13:45:00Z',
    updatedAt: '2024-02-07T11:15:00Z',
    content: `
CREATIVE BRIEF

PROJECT OVERVIEW:
Project Name: [Campaign/Project Name]
Client/Brand: [Brand Name]
Project Type: ☐ Campaign ☐ Ad ☐ Video ☐ Website ☐ Print ☐ Social Media ☐ Other: [Specify]
Date: [Date]
Project Owner: [Name]
Creative Lead: [Name]

PROJECT DETAILS:
Brief Number: [#]
Status: ☐ Draft ☐ In Review ☐ Approved ☐ In Production
Priority: ☐ High ☐ Medium ☐ Low

BACKGROUND & CONTEXT:

Company/Brand Overview:
[Brief description of the company, brand position, and relevant history]

Current Situation:
[What's happening now that makes this project necessary?]

The Problem/Opportunity:
[What challenge are we solving or what opportunity are we capitalizing on?]

Why Now:
[Why is this the right time for this creative project?]

---

PROJECT OBJECTIVES:

Primary Objective:
[The single most important thing this project needs to achieve]

Secondary Objectives:
• [Objective 1]
• [Objective 2]
• [Objective 3]

Success Criteria:
How we'll know this project succeeded:
• [Metric 1]: [Target]
• [Metric 2]: [Target]
• [Metric 3]: [Target]

---

TARGET AUDIENCE:

Primary Audience:
Who: [Demographic and psychographic description]
Age: [Range]
Gender: [If relevant]
Location: [Geographic]
Occupation/Industry: [If relevant]
Income: [If relevant]

Mindset & Behavior:
• Current beliefs: [What they think now]
• Attitudes: [How they feel]
• Behaviors: [What they do]
• Pain points: [What frustrates them]
• Desires: [What they want]

Relationship with Our Brand:
☐ Unaware
☐ Aware but never used
☐ Current customer
☐ Lapsed customer
☐ Competitor's customer

Media Consumption:
Where they spend time:
• [Platform/Channel 1]
• [Platform/Channel 2]
• [Platform/Channel 3]

Secondary Audience (if applicable):
[Brief description]

---

DESIRED AUDIENCE RESPONSE:

What We Want Them to Think:
[The perception or belief we want to create]

What We Want Them to Feel:
[The emotional response we're aiming for]

What We Want Them to Do:
[The specific action we want them to take]

---

KEY MESSAGE:

Single Most Important Thing to Communicate:
[One clear, concise message - if they remember only one thing, what should it be?]

Supporting Messages:
• [Supporting point 1]
• [Supporting point 2]
• [Supporting point 3]

Message Tone:
☐ Professional ☐ Casual ☐ Friendly ☐ Authoritative ☐ Playful ☐ Serious
☐ Inspiring ☐ Urgent ☐ Reassuring ☐ Educational ☐ Entertaining

Brand Voice:
How we should sound:
• [Voice characteristic 1 - e.g., "Confident but not arrogant"]
• [Voice characteristic 2 - e.g., "Simple but not simplistic"]
• [Voice characteristic 3 - e.g., "Warm but professional"]

Words/Phrases to Use:
• [Word/phrase 1]
• [Word/phrase 2]
• [Word/phrase 3]

Words/Phrases to Avoid:
• [Word/phrase 1]
• [Word/phrase 2]

---

UNIQUE SELLING PROPOSITION (USP):

What Makes Us Different:
[What sets us apart from competitors]

Reason to Believe:
[Proof points that support our claims]
• [Proof point 1]
• [Proof point 2]
• [Proof point 3]

Competitive Advantage:
[Why should they choose us over alternatives?]

---

CALL-TO-ACTION (CTA):

Primary CTA:
[Specific action we want them to take]

CTA Copy:
[Exact wording of the CTA]

Secondary CTA (if applicable):
[Alternative or next-step action]

---

MANDATORIES:

Must-Include Elements:
• [Mandatory element 1 - e.g., "Company logo"]
• [Mandatory element 2 - e.g., "Legal disclaimer"]
• [Mandatory element 3 - e.g., "Specific tagline"]
• [Mandatory element 4]

Brand Guidelines:
☐ Follow existing brand guidelines strictly
☐ Can push creative boundaries within brand framework
☐ This is a brand evolution - new direction expected

Logo Usage:
[Specifications for logo placement, size, color variations]

Color Palette:
Primary Colors: [Colors with hex codes]
Secondary Colors: [Colors with hex codes]
Accent Colors: [Colors with hex codes]

Typography:
Headlines: [Font name and specifications]
Body Copy: [Font name and specifications]
Captions: [Font name and specifications]

Imagery Style:
☐ Photography ☐ Illustration ☐ 3D Renders ☐ Icons ☐ Mixed Media
Style: [Description of desired visual style]

Legal/Compliance Requirements:
• [Requirement 1]
• [Requirement 2]
• [Requirement 3]

---

CREATIVE CONSIDERATIONS:

Tone & Mood:
[Description of the overall feeling the creative should evoke]

Creative Direction/Style:
[Guidance on the creative approach - modern, minimalist, bold, etc.]

Inspiration/References:
Examples of creative we like (and why):
• [Example 1]: [What we like about it]
• [Example 2]: [What we like about it]
• [Example 3]: [What we like about it]

Examples to Avoid:
• [Example 1]: [Why to avoid]
• [Example 2]: [Why to avoid]

Cultural Considerations:
[Any cultural sensitivities or considerations to keep in mind]

Accessibility Requirements:
☐ WCAG 2.1 AA compliant
☐ Closed captions for video
☐ Alt text for images
☐ Color contrast requirements
☐ Other: [Specify]

---

DELIVERABLES:

What We Need:
Deliverable | Format | Dimensions/Specs | Quantity | Due Date
------------|--------|------------------|----------|----------
[Item 1] | [File type] | [Specs] | [#] | [Date]
[Item 2] | [File type] | [Specs] | [#] | [Date]
[Item 3] | [File type] | [Specs] | [#] | [Date]
[Item 4] | [File type] | [Specs] | [#] | [Date]

Examples:
• Print ad - PDF, 8.5x11", 300 DPI, CMYK - Due: [Date]
• Social media graphics - PNG, 1080x1080px, 72 DPI, RGB - Due: [Date]
• Video - MP4, 1920x1080, 30fps, under 2min - Due: [Date]

File Naming Convention:
[Specific naming structure to use]

Where to Deliver Files:
[Platform, folder location, or email]

---

TIMELINE & MILESTONES:

Project Kick-off: [Date]
Initial Concepts Due: [Date]
Internal Review: [Date]
Revisions Due: [Date]
Final Approval: [Date]
Final Delivery: [Date]
Launch/Go-Live: [Date]

Review Process:
Round 1 Reviews: [Who reviews and by when]
Round 2 Reviews: [Who reviews and by when]
Final Sign-off: [Who gives final approval]

---

BUDGET:

Total Budget: $[Amount]

Budget Breakdown:
• Creative development: $[Amount]
• Production: $[Amount]
• Media/Placement: $[Amount]
• Contingency: $[Amount]

Budget Notes:
[Any important context about the budget]

---

DISTRIBUTION & PLACEMENT:

Where This Will Live:
☐ Website: [Specific pages]
☐ Social Media: [Which platforms]
☐ Email: [Which campaigns]
☐ Print: [Which publications]
☐ Out-of-home: [Where]
☐ Video: [Which channels]
☐ Events: [Which events]
☐ Other: [Specify]

Media Specs for Each Placement:
[Specific technical requirements for each channel]

Campaign Duration:
Start Date: [Date]
End Date: [Date]
Flight Dates: [If multiple waves]

---

STAKEHOLDERS:

Project Team:
• Project Owner: [Name]
• Creative Director: [Name]
• Copywriter: [Name]
• Art Director/Designer: [Name]
• Producer (if video): [Name]
• Developer (if digital): [Name]
• Project Manager: [Name]

Approvers:
• [Role]: [Name]
• [Role]: [Name]
• [Role]: [Name]

Who Needs to be Involved:
• [Department/Person 1]: [When/Why]
• [Department/Person 2]: [When/Why]

---

COMPETITIVE LANDSCAPE:

Main Competitors:
Competitor 1: [Name]
Their Approach: [How they're messaging/positioning]
What They Do Well: [Strengths]
Opportunity for Us: [How we can differentiate]

Competitor 2: [Name]
Their Approach: [How they're messaging/positioning]
What They Do Well: [Strengths]
Opportunity for Us: [How we can differentiate]

Competitor 3: [Name]
Their Approach: [How they're messaging/positioning]
What They Do Well: [Strengths]
Opportunity for Us: [How we can differentiate]

---

MEASUREMENT & SUCCESS:

How We'll Measure Success:

Awareness Metrics:
• [Metric 1]: [Target]
• [Metric 2]: [Target]

Engagement Metrics:
• [Metric 1]: [Target]
• [Metric 2]: [Target]

Conversion Metrics:
• [Metric 1]: [Target]
• [Metric 2]: [Target]

Tracking Plan:
[How we'll track and measure performance]

Reporting Frequency:
[How often we'll review performance]

---

CHALLENGES & CONSTRAINTS:

Known Challenges:
• [Challenge 1]
• [Challenge 2]
• [Challenge 3]

Constraints:
• Time: [Time constraints]
• Budget: [Budget limitations]
• Resources: [Resource constraints]
• Technical: [Technical limitations]
• Legal: [Legal restrictions]

Potential Risks:
• [Risk 1]: [Mitigation plan]
• [Risk 2]: [Mitigation plan]

---

ADDITIONAL INFORMATION:

Related Projects/Campaigns:
• [Project 1]: [How it relates]
• [Project 2]: [How it relates]

Assets Available:
• [Asset 1]: [Where to find it]
• [Asset 2]: [Where to find it]
• [Asset 3]: [Where to find it]

Research/Data:
• [Research finding 1]
• [Research finding 2]
• [Research finding 3]

---

QUESTIONS FOR THE CREATIVE TEAM:

[List any open questions that need to be answered]
• [Question 1]
• [Question 2]
• [Question 3]

---

APPROVALS:

Creative Brief Prepared by: _________________ Date: _______
[Name, Title]

Reviewed by: _________________ Date: _______
[Name, Title]

Approved by: _________________ Date: _______
[Name, Title]

Creative Team Lead: _________________ Date: _______
[Name, Title]

---

CREATIVE TEAM NOTES:
[Space for creative team to add notes, ideas, or questions during the creative development process]

---

BRIEF SUMMARY (One-Pager):

For quick reference, answer these questions:

What are we making? [One sentence]
Who is it for? [One sentence]
What do we want them to do? [One sentence]
What's the one thing we want them to remember? [One sentence]
Why should they care? [One sentence]
When do we need it? [Date]
    `
  },
  {
    id: 'marketing-press-release',
    title: 'Press Release Template',
    description: 'Professional press release template following standard PR format and best practices',
    category: 'marketing',
    downloadCount: 4670,
    fileSize: '26.3 KB',
    rating: 4.6,
    tags: ['press-release', 'pr', 'media', 'announcement', 'news'],
    createdAt: '2024-01-25T15:00:00Z',
    updatedAt: '2024-02-08T10:00:00Z',
    content: `
PRESS RELEASE

FOR IMMEDIATE RELEASE
[or: EMBARGOED UNTIL [Date and Time]]

[HEADLINE IN TITLE CASE - Grab Attention, Around 10-15 Words]
[Make it newsworthy, specific, and compelling]

[SUBHEADLINE (Optional) - Additional Context or Detail]
[Expand on the headline with more specific information]

[CITY, STATE] – [Date] – [Opening paragraph answering Who, What, When, Where, Why, and How]

[The first paragraph should be 2-4 sentences that summarize the entire announcement. This should contain all the essential information in case someone only reads this paragraph.]

[COMPANY NAME], [brief description of company], today announced [the news]. [Key detail about significance or impact]. [Quote-worthy fact or statistic if available].

[Second Paragraph - Expand on the News]

[Provide more details about the announcement. Include relevant facts, figures, and context. Answer questions a reader might have. This paragraph should elaborate on the WHO and WHAT from the first paragraph.]

[Third Paragraph - Quote from Company Executive]

"[Quote from CEO, President, or relevant executive - 2-3 sentences maximum]," said [Full Name], [Title] of [Company Name]. "[The quote should provide insight, perspective, or emotional context that can't be communicated through straight facts. It should sound natural and conversational, not like marketing copy.]"

[Fourth Paragraph - Additional Details and Benefits]

[Provide more context, details about benefits to customers, users, or the industry. Include specific examples of how this news matters. If announcing a product, describe key features and advantages. If announcing a partnership, explain what each party brings to the table.]

[Fifth Paragraph - Quote from Partner, Customer, or Industry Expert (Optional)]

"[Quote from relevant third party that adds credibility and outside perspective]," said [Full Name], [Title] of [Company/Organization Name]. "[This quote should validate the importance of the announcement and provide an external viewpoint.]"

[Sixth Paragraph - Background and Context]

[Provide additional context, background information, or the bigger picture. Explain how this fits into company strategy, industry trends, or market needs. Include any relevant statistics or market research if applicable.]

[Seventh Paragraph - Availability and Next Steps]

[Information about when/how the product/service/initiative will be available. Include pricing if relevant and approved. Provide clear next steps for interested parties.]

[For more information about [topic], visit [URL] or contact [contact information].]

[Closing Paragraph - Company Boilerplate]

About [Company Name]
[Standard company description, usually 3-5 sentences. This should be the same across all press releases unless company positioning changes. Include:
- What the company does
- Who they serve
- Key differentiators or achievements
- Notable clients or metrics if impressive
- Where they're based
End with company URL]

[If Multiple Companies Involved - Additional Boilerplate]

About [Partner Company Name]
[Same format as above]

Media Contact:
[Full Name]
[Title]
[Company Name]
Phone: [(XXX) XXX-XXXX]
Email: [email@company.com]

[If using PR agency]
PR Contact:
[Agency Name]
[Contact Person]
Phone: [(XXX) XXX-XXXX]
Email: [email@agency.com]

###
[The "###" or "-30-" marks the end of the press release]

---

PRESS RELEASE WRITING GUIDE:

HEADLINE BEST PRACTICES:
☐ Make it newsworthy, not promotional
☐ Include company name (usually)
☐ Use active voice
☐ Keep under 100 characters if possible
☐ Include key keywords for SEO
☐ Front-load the most important information

Good Headline Examples:
• [Company Name] Launches [Product] to Solve [Problem] for [Target Audience]
• [Company Name] Secures $[X]M in Series [X] Funding Led by [Investor]
• [Company Name] Partners with [Partner] to [Outcome]
• [Company Name] Expands to [New Market/Location]

LEAD PARAGRAPH CHECKLIST:
☐ City and date included
☐ Company name included
☐ Announcement clearly stated in first sentence
☐ Key facts included (who, what, when, where, why)
☐ Under 50 words total
☐ Can stand alone as a complete story

QUOTE GUIDELINES:
☐ Sounds like a real person, not marketing copy
☐ Adds perspective or emotion, not just facts
☐ 2-3 sentences maximum
☐ Includes attribution with full name and title
☐ Provides insight that can't be conveyed through facts alone

THINGS TO AVOID:
• Excessive superlatives ("revolutionary," "amazing," "best ever")
• Marketing jargon and buzzwords
• First-person pronouns (we, our) outside of quotes
• Passive voice
• Burying the lead
• Making readers work to understand the news
• Including unverified claims or hyperbole
• Forgetting contact information

FORMATTING STANDARDS:
☐ Use standard press release format
☐ Double-space between paragraphs
☐ Use clear, readable font (Times New Roman or Arial, 12pt)
☐ Keep to one or two pages maximum
☐ Use proper AP Style
☐ Include release date at top
☐ End with ### or -30-
☐ Include page numbers if multiple pages ("Page X of Y")

LENGTH GUIDELINES:
Ideal length: 400-600 words (about 1-1.5 pages)
Maximum: 800 words
Minimum: 300 words

TYPES OF PRESS RELEASES (Choose the Right Format):

New Product/Service Launch:
• Focus on what problem it solves
• Include availability and pricing
• Highlight key features and benefits
• Include customer quote if possible

Company Milestone:
• Put the achievement in context
• Include relevant statistics
• Explain significance to industry
• Look forward to future plans

Partnership/Collaboration:
• Explain what each party brings
• Highlight mutual benefits
• Include quotes from both parties
• Describe expected outcomes

Funding Announcement:
• State amount and funding round
• Name lead investors
• Explain what funds will be used for
• Include investor quote

Executive Appointment:
• Include bio of new executive
• Explain their background and qualifications
• Quote from CEO about appointment
• Quote from new executive

Event Announcement:
• Include who, what, when, where
• Explain who should attend and why
• Provide registration information
• Include speaker information if relevant

Research/Study Results:
• Lead with most surprising/newsworthy finding
• Include methodology details
• Provide context and implications
• Include quotes interpreting the data

Award/Recognition:
• Name the award and presenting organization
• Explain the significance
• Include selection criteria
• Thank relevant parties in quote

DISTRIBUTION CHECKLIST:

Before Sending:
☐ Proofread multiple times
☐ Check all links work
☐ Verify all facts and figures
☐ Get legal approval if needed
☐ Get executive approval on quotes
☐ Test contact information
☐ Optimize headline for SEO
☐ Create social media versions
☐ Prepare media kit if needed

Distribution Plan:
☐ Primary media list identified
☐ Trade publications identified
☐ Local media identified (if relevant)
☐ Industry influencers identified
☐ Distribution service selected (if using)
☐ Timing optimized (Tuesday-Thursday, 10am-2pm best)
☐ Follow-up plan created

Post-Distribution:
☐ Post to company website/newsroom
☐ Share on social media
☐ Email to stakeholders
☐ Follow up with key journalists
☐ Monitor coverage
☐ Track metrics (pickups, reach, etc.)
☐ Update media database based on responses

SEO OPTIMIZATION:

Keywords to Include:
• [Primary keyword]
• [Secondary keyword]
• [Long-tail keyword]

Meta Description (if posting online):
[150-160 character description with keywords]

Multimedia to Include:
☐ High-resolution images (300 DPI for print, 72 DPI for web)
☐ Infographics
☐ Video (if applicable)
☐ Logos (multiple formats)
☐ Executive headshots
☐ Product photos

MEDIA KIT CHECKLIST (If Creating):
☐ Press release
☐ Company backgrounder
☐ Executive bios
☐ High-res images
☐ Company logo (various formats)
☐ Fact sheet
☐ FAQs
☐ Product specifications
☐ Demo video (if applicable)
☐ Contact information

MEASUREMENT:

Success Metrics:
• Media pickups: [Target number]
• Tier 1 publications: [Target number]
• Social media shares: [Target number]
• Website traffic from release: [Target number]
• Backlinks generated: [Target number]
• Estimated reach: [Target number]
• Estimated ad value equivalency: $[Amount]

LEGAL DISCLAIMER (If Required):
[Standard legal language about forward-looking statements, etc.]

APPROVALS:

Drafted by: _________________ Date: _______
[PR Manager/Writer]

Reviewed by: _________________ Date: _______
[Marketing Director]

Legal Approved: _________________ Date: _______
[Legal Team]

Executive Approved: _________________ Date: _______
[CEO/Authorized Executive]

DISTRIBUTION TRACKING:

Distribution Date: _______
Distribution Method: [Wire service, Direct email, Both]
Number of Media Outlets Contacted: [#]
Pickups Received: [#]
Top Placements: [List key media outlets that covered it]
    `
  }
];

// Add marketing templates to the main templates array
templates.push(...marketingTemplates);

// Legal & Admin Templates
const legalAdminTemplates: DocumentTemplate[] = [
  {
    id: 'legal-service-agreement',
    title: 'Service Agreement',
    description: 'Comprehensive service agreement template for defining terms between service provider and client',
    category: 'legal',
    downloadCount: 8920,
    fileSize: '52.3 KB',
    rating: 4.8,
    tags: ['service-agreement', 'contract', 'legal', 'business', 'terms'],
    createdAt: '2024-01-08T10:00:00Z',
    updatedAt: '2024-01-29T16:45:00Z',
    content: `
SERVICE AGREEMENT

This Service Agreement ("Agreement") is entered into as of [Date] ("Effective Date") by and between:

SERVICE PROVIDER:
[Company/Individual Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]

CLIENT:
[Company/Individual Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]

WHEREAS, Provider offers [description of services]; and
WHEREAS, Client desires to engage Provider to perform such services;

NOW, THEREFORE, in consideration of the mutual covenants and agreements herein contained, the parties agree as follows:

1. SERVICES

1.1 Scope of Services
Provider agrees to provide the following services ("Services"):
• [Service 1]
• [Service 2]
• [Service 3]
• [Additional services as specified in Exhibit A]

1.2 Service Specifications
• Deliverables: [List specific deliverables]
• Performance Standards: [Quality standards, response times, etc.]
• Service Location: [On-site, remote, hybrid]
• Service Hours: [Hours of availability]

2. TERM AND TERMINATION

2.1 Term
This Agreement shall commence on [Start Date] and continue until [End Date], unless terminated earlier in accordance with this Agreement.

2.2 Termination for Convenience
Either party may terminate this Agreement with [#] days written notice.

2.3 Termination for Cause
Either party may terminate immediately upon written notice if the other party:
• Materially breaches this Agreement and fails to cure within [#] days
• Becomes insolvent or files for bankruptcy
• Engages in fraud or illegal activities

2.4 Effect of Termination
Upon termination:
• Client shall pay for all services rendered through the termination date
• Provider shall deliver all work product and Client property
• All confidential information shall be returned or destroyed
• Sections [#] shall survive termination

3. COMPENSATION AND PAYMENT

3.1 Fees
Client agrees to pay Provider as follows:
☐ Fixed Fee: $[Amount] for [scope]
☐ Hourly Rate: $[Rate] per hour, estimated [#] hours
☐ Monthly Retainer: $[Amount] per month
☐ Other: [Specify payment structure]

3.2 Expenses
☐ Included in fees
☐ Reimbursable with prior approval (estimated $[Amount])
☐ Subject to markup of [%]

3.3 Payment Terms
• Invoice Schedule: [Monthly, upon milestones, etc.]
• Payment Due: Net [#] days from invoice date
• Payment Method: [Wire transfer, check, ACH, etc.]
• Late Fee: [%] per month on overdue amounts

3.4 Additional Services
Services beyond the scope shall be billed at:
• Hourly Rate: $[Rate]
• Or as mutually agreed in writing

4. RESPONSIBILITIES

4.1 Provider Responsibilities
Provider shall:
• Perform Services in a professional and workmanlike manner
• Comply with all applicable laws and regulations
• Provide qualified personnel
• Maintain adequate insurance coverage
• [Additional responsibilities]

4.2 Client Responsibilities
Client shall:
• Provide timely access to information, systems, and personnel
• Review and approve deliverables within [#] business days
• Make timely payments
• Provide workspace/equipment if applicable: [Details]
• [Additional responsibilities]

5. INTELLECTUAL PROPERTY

5.1 Ownership of Deliverables
☐ Client owns all work product upon full payment
☐ Provider retains ownership; Client receives license: [Type of license]
☐ Joint ownership as specified: [Details]

5.2 Pre-Existing Materials
Provider retains ownership of pre-existing:
• Methodologies, tools, and templates
• Generic knowledge and experience
• [Other pre-existing IP]

5.3 License Grant
[If applicable] Provider grants Client a [exclusive/non-exclusive], [perpetual/term-limited], [transferable/non-transferable] license to use deliverables for [permitted uses].

6. CONFIDENTIALITY

6.1 Confidential Information
Each party agrees to maintain in confidence all proprietary information of the other party, including:
• Business strategies and financial information
• Technical data and trade secrets
• Customer and supplier information
• This Agreement and its terms

6.2 Obligations
Receiving party shall:
• Use Confidential Information solely for this Agreement
• Protect it with at least the same care as its own confidential information
• Limit disclosure to employees/contractors with need to know
• Not disclose to third parties without prior written consent

6.3 Exceptions
Confidentiality obligations do not apply to information that:
• Is publicly available through no breach of this Agreement
• Was rightfully known prior to disclosure
• Is independently developed
• Must be disclosed by law (with prompt notice to disclosing party)

6.4 Duration
Confidentiality obligations survive for [#] years after termination.

7. WARRANTIES AND REPRESENTATIONS

7.1 Provider Warranties
Provider represents and warrants that:
• It has the right and authority to enter this Agreement
• Services will be performed in a professional manner
• Services will conform to specifications and industry standards
• Work product will not infringe third-party rights
• It maintains appropriate licenses and qualifications

7.2 Client Warranties
Client represents and warrants that:
• It has authority to enter this Agreement
• Materials provided do not infringe third-party rights
• It will comply with applicable laws

7.3 Warranty Period
Warranty period for deliverables: [#] days from acceptance

7.4 Warranty Remedy
Exclusive remedy: Provider will re-perform non-conforming Services at no charge

7.5 DISCLAIMER
EXCEPT AS EXPRESSLY PROVIDED HEREIN, SERVICES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.

8. LIMITATION OF LIABILITY

8.1 Cap on Liability
EXCEPT FOR EXCLUDED CLAIMS, EACH PARTY'S TOTAL LIABILITY SHALL NOT EXCEED:
☐ The fees paid in the [12] months preceding the claim
☐ $[Amount]
☐ [Other limitation]

8.2 Excluded Claims
Limitations do not apply to:
• Gross negligence or willful misconduct
• Breach of confidentiality
• Infringement of intellectual property
• Death or bodily injury
• [Other exclusions]

8.3 CONSEQUENTIAL DAMAGES
IN NO EVENT SHALL EITHER PARTY BE LIABLE FOR INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, OR PUNITIVE DAMAGES, INCLUDING LOST PROFITS, EVEN IF ADVISED OF THEIR POSSIBILITY.

9. INDEMNIFICATION

9.1 Provider Indemnification
Provider shall indemnify and defend Client against claims arising from:
• Provider's negligence or willful misconduct
• Infringement of third-party IP rights by deliverables
• Breach of Provider's representations and warranties

9.2 Client Indemnification
Client shall indemnify and defend Provider against claims arising from:
• Client's negligence or willful misconduct
• Client's use of deliverables beyond scope of license
• Materials or instructions provided by Client

9.3 Indemnification Procedure
Indemnified party must:
• Promptly notify indemnifying party of claim
• Allow indemnifying party to control defense
• Cooperate reasonably in defense

10. INDEPENDENT CONTRACTOR

10.1 Relationship
Provider is an independent contractor, not an employee, partner, or agent of Client.

10.2 Taxes and Benefits
Provider is responsible for:
• All taxes (income, self-employment, etc.)
• Insurance (liability, workers' compensation, etc.)
• Benefits
• Compliance with labor and employment laws

10.3 No Benefits
Provider is not entitled to employee benefits including health insurance, paid leave, or retirement benefits.

11. INSURANCE

Provider shall maintain the following insurance:
• General Liability: $[Amount] per occurrence
• Professional Liability (E&O): $[Amount] per claim
• Workers' Compensation: As required by law
• Cyber Liability: $[Amount] (if applicable)

Certificates of insurance shall be provided upon request.

12. GENERAL PROVISIONS

12.1 Entire Agreement
This Agreement constitutes the entire agreement and supersedes all prior agreements and understandings.

12.2 Amendments
Amendments must be in writing and signed by both parties.

12.3 Assignment
Neither party may assign this Agreement without prior written consent, except to a successor in a merger or acquisition.

12.4 Notices
All notices shall be in writing and delivered to:

Provider: [Contact information]
Client: [Contact information]

12.5 Governing Law
This Agreement shall be governed by the laws of [State/Country], without regard to conflicts of law principles.

12.6 Dispute Resolution
☐ Mediation required before litigation
☐ Binding arbitration under [AAA/JAMS] rules in [Location]
☐ Litigation in courts of [Jurisdiction]

12.7 Attorney's Fees
Prevailing party in dispute shall be entitled to reasonable attorney's fees and costs.

12.8 Severability
If any provision is invalid, the remainder shall continue in effect.

12.9 Waiver
Failure to enforce any provision does not waive the right to enforce it later.

12.10 Force Majeure
Neither party is liable for delays caused by events beyond reasonable control (acts of God, war, strikes, pandemics, etc.).

12.11 Publicity
☐ Provider may list Client as reference and use in marketing materials
☐ No publicity without prior written approval

12.12 Counterparts
This Agreement may be executed in counterparts and by electronic signature.

13. SPECIAL PROVISIONS

[Any industry-specific or custom provisions]


IN WITNESS WHEREOF, the parties have executed this Agreement as of the Effective Date.

SERVICE PROVIDER:

Signature: _________________________
Name: [Printed Name]
Title: [Title]
Date: _____________________________

CLIENT:

Signature: _________________________
Name: [Printed Name]
Title: [Title]
Date: _____________________________


EXHIBIT A: DETAILED SCOPE OF SERVICES

[Detailed description of services, deliverables, milestones, specifications, acceptance criteria, etc.]

EXHIBIT B: PAYMENT SCHEDULE

[Detailed payment milestones and amounts]
    `
  },
  {
    id: 'legal-partnership-agreement',
    title: 'Partnership Agreement',
    description: 'Detailed partnership agreement for establishing business partnerships with profit sharing and governance terms',
    category: 'legal',
    downloadCount: 7340,
    fileSize: '58.7 KB',
    rating: 4.7,
    tags: ['partnership', 'business', 'legal', 'agreement', 'equity'],
    createdAt: '2024-01-09T11:30:00Z',
    updatedAt: '2024-01-30T10:20:00Z',
    content: `
PARTNERSHIP AGREEMENT

This Partnership Agreement ("Agreement") is entered into as of [Date] ("Effective Date") by and among the following partners ("Partners"):

PARTNER 1:
[Full Legal Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]
Ownership Percentage: [%]

PARTNER 2:
[Full Legal Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]
Ownership Percentage: [%]

[Additional Partners]

RECITALS

WHEREAS, the Partners wish to form a partnership for the purpose of [business purpose];
WHEREAS, the Partners desire to set forth the terms and conditions of their partnership;

NOW, THEREFORE, in consideration of the mutual covenants herein, the Partners agree:

1. FORMATION AND NAME

1.1 Partnership Name
The name of the partnership shall be: [Partnership Name] ("Partnership")

1.2 Principal Place of Business
[Address]
[City, State, ZIP Code]

1.3 Type of Partnership
☐ General Partnership
☐ Limited Partnership
☐ Limited Liability Partnership (LLP)

1.4 Formation
The Partnership is formed under the laws of [State], effective [Date].

2. PURPOSE AND BUSINESS

2.1 Business Purpose
The Partnership is formed for the following purposes:
• [Primary business activity]
• [Secondary activities]
• [Other permitted activities]

2.2 Scope of Business
The Partnership may engage in any lawful business activities agreed upon by the Partners.

2.3 Term
The Partnership shall continue until [Date] or until dissolved according to this Agreement.

3. CAPITAL CONTRIBUTIONS

3.1 Initial Contributions
Each Partner shall contribute the following:

Partner 1: [Name]
• Cash: $[Amount]
• Property: [Description, valued at $Amount]
• Services: [Description, valued at $Amount]
• Total: $[Amount]

Partner 2: [Name]
• Cash: $[Amount]
• Property: [Description, valued at $Amount]
• Services: [Description, valued at $Amount]
• Total: $[Amount]

[Additional Partners]

3.2 Additional Contributions
• Required contributions may be called by [majority/unanimous] vote
• Notice period: [#] days
• Failure to contribute: [Consequences - dilution, forced sale, etc.]

3.3 Return of Contributions
Capital contributions shall not be returned except upon dissolution or as specifically provided herein.

3.4 Interest on Capital
Partners shall [receive/not receive] interest on capital contributions at [%] per annum.

4. OWNERSHIP AND PROFIT SHARING

4.1 Ownership Percentages
Partner 1: [Name] - [%]
Partner 2: [Name] - [%]
[Additional Partners]
Total: 100%

4.2 Profit and Loss Allocation
Profits and losses shall be allocated according to ownership percentages, unless otherwise agreed.

4.3 Distributions
• Timing: [Quarterly, annually, as determined by Partners]
• Amount: As determined by [majority/unanimous] vote
• Minimum distribution: [If applicable]
• Tax distributions: Sufficient to cover each Partner's tax liability

4.4 Draws
Each Partner may take monthly draws of up to $[Amount], subject to available cash flow.

5. MANAGEMENT AND DECISION MAKING

5.1 Management Structure
☐ All Partners manage equally
☐ Managing Partner(s): [Name(s)]
☐ Management Committee: [Composition]

5.2 Voting Rights
• Routine matters: [Majority/Unanimous] vote required
• Major decisions: [Unanimous/Supermajority] vote required

5.3 Major Decisions Requiring [Unanimous/Supermajority] Approval
• Admission of new Partners
• Sale, merger, or dissolution of Partnership
• Purchase or sale of Partnership assets exceeding $[Amount]
• Borrowing exceeding $[Amount]
• Execution of leases or contracts exceeding [$ or term]
• Amendment of this Agreement
• Changes to capital structure
• [Other major decisions]

5.4 Authority of Partners
Each Partner has authority to:
• Bind the Partnership in ordinary course of business
• Sign contracts up to $[Amount]
• [Other routine authorities]

Each Partner does NOT have authority to:
• Sell Partnership assets without approval
• Admit new Partners
• Settle litigation exceeding $[Amount]
• [Other restricted activities]

5.5 Meetings
• Regular meetings: [Frequency]
• Special meetings: Called by any Partner with [#] days notice
• Quorum: [#] Partners or [%] of ownership
• Action by written consent permitted

6. DUTIES AND RESPONSIBILITIES

6.1 Partner Duties
Each Partner shall:
• Devote [full-time/[#] hours per week] to Partnership business
• Act in good faith and in Partnership's best interests
• Comply with all applicable laws and regulations
• Maintain accurate books and records
• [Specific duties by Partner]

Partner 1: [Name]
• [Specific responsibilities]

Partner 2: [Name]
• [Specific responsibilities]

6.2 Fiduciary Duties
Partners owe each other and the Partnership duties of loyalty, care, and good faith.

6.3 Confidentiality
Partners shall maintain confidentiality of Partnership information during and after the Partnership.

6.4 Non-Compete
During the term of this Agreement and for [#] years thereafter, Partners shall not:
• Compete directly with the Partnership in [geographic area]
• Solicit Partnership customers or employees
• [Other non-compete restrictions]

Exception: [If applicable - e.g., approved side businesses]

6.5 Compensation
Beyond profit distributions, Partners may receive:
• Salary: [Details by Partner]
• Bonuses: [Criteria]
• Benefits: [Health insurance, retirement, etc.]
• Expense reimbursement: [Policy]

7. BOOKS, RECORDS, AND ACCOUNTING

7.1 Fiscal Year
The Partnership's fiscal year shall be [Calendar year / Other period].

7.2 Accounting Method
The Partnership shall use the [cash/accrual] method of accounting.

7.3 Books and Records
The Partnership shall maintain complete books and records at the principal place of business.

7.4 Bank Accounts
• Bank: [Name of bank]
• Signatories: [Names and any limitations]
• Approval required for withdrawals exceeding: $[Amount]

7.5 Financial Statements
• Annual financial statements shall be prepared by [Date]
• Prepared by: ☐ Internal ☐ CPA firm
• Audit required: ☐ Yes ☐ No

7.6 Tax Returns
• Partnership tax returns (Form 1065) shall be filed timely
• K-1s distributed to Partners by [Date]
• Tax matters partner: [Name]

7.7 Access to Records
Partners have full access to all books and records at reasonable times.

8. TRANSFER OF PARTNERSHIP INTERESTS

8.1 Restrictions on Transfer
Partnership interests may not be transferred without prior written consent of [all other Partners / majority].

8.2 Right of First Refusal
Before transferring any interest, the selling Partner must:
• Provide written notice with terms
• Other Partners have [#] days to purchase on same terms
• Pro-rata allocation among purchasing Partners

8.3 Permitted Transfers
Without consent, Partners may transfer to:
☐ Spouse or children
☐ Trust for estate planning
☐ [Other permitted transferees]

8.4 Buy-Sell Provisions
Upon occurrence of trigger events, Partnership interests shall be purchased:

Trigger Events:
• Death
• Disability (unable to work for [#] months)
• Bankruptcy
• Divorce (if interest at risk)
• Retirement (age [#] or after [#] years)
• Voluntary withdrawal
• Termination for cause

8.5 Valuation Method
☐ Formula: [e.g., Book value, Multiple of earnings]
☐ Independent appraisal by [Appraiser]
☐ Agreement of Partners
☐ Annual valuation: Last valuation was $[Amount] on [Date]

8.6 Payment Terms
• Down payment: [%] within [#] days
• Balance: Paid over [#] years
• Interest rate: [%] per annum
• Security: [If applicable]

8.7 Funding Mechanism
Purchase shall be funded by:
☐ Partnership redemption
☐ Cross-purchase by other Partners
☐ Life insurance (for death)
☐ Disability insurance (for disability)

9. ADMISSION OF NEW PARTNERS

9.1 Approval Required
Admission requires [unanimous/supermajority] approval of existing Partners.

9.2 Terms of Admission
New Partner must:
• Make capital contribution of $[Amount] or [% of valuation]
• Execute counterpart of this Agreement
• [Other requirements]

9.3 Effect on Ownership
Existing Partners' interests shall be [diluted proportionally / protected].

10. WITHDRAWAL AND RETIREMENT

10.1 Voluntary Withdrawal
A Partner may withdraw upon [#] months written notice.

10.2 Effect of Withdrawal
Upon withdrawal:
• Withdrawing Partner's interest shall be purchased per Section 8
• Partner loses all management rights immediately
• Non-compete provisions remain in effect
• [Other consequences]

11. DISSOLUTION AND WINDING UP

11.1 Events Causing Dissolution
The Partnership shall dissolve upon:
• Agreement of Partners holding [%] of interests
• Expiration of term
• [Number] or fewer Partners remain
• Bankruptcy of Partnership
• Court order

11.2 Winding Up Process
Upon dissolution:
• Business operations cease except as necessary for winding up
• Liquidating Partner(s): [Name or selection method]
• Assets shall be sold or distributed in kind
• Creditors paid in order of priority
• Remaining assets distributed to Partners per ownership %

11.3 Order of Distribution
1. Creditors (including Partner loans)
2. Return of capital contributions
3. Remaining profits per ownership percentages

12. DISPUTE RESOLUTION

12.1 Good Faith Negotiation
Partners shall first attempt to resolve disputes through good faith negotiation.

12.2 Mediation
If negotiation fails, disputes shall be submitted to mediation before [Organization] in [Location].

12.3 Arbitration
☐ Binding arbitration required under [AAA/JAMS] rules
☐ Arbitration optional
☐ Litigation permitted after mediation

12.4 Deadlock
If Partners are deadlocked on major decision:
☐ Mediation required
☐ Buy-sell triggered: [Details]
☐ [Other deadlock mechanism]

13. GENERAL PROVISIONS

13.1 Entire Agreement
This Agreement constitutes the entire agreement among the Partners.

13.2 Amendments
Amendments require [written consent of all Partners / supermajority].

13.3 Governing Law
This Agreement shall be governed by the laws of [State].

13.4 Notices
Notices shall be sent to addresses listed above or as updated in writing.

13.5 Severability
Invalid provisions shall be severed; remainder continues in effect.

13.6 Waiver
Waiver of any provision must be in writing and does not constitute ongoing waiver.

13.7 Successors and Assigns
This Agreement binds and benefits Partners' successors and permitted assigns.

13.8 Counterparts
This Agreement may be executed in counterparts.

14. SPECIAL PROVISIONS

[Any additional provisions specific to the partnership]


IN WITNESS WHEREOF, the Partners have executed this Agreement as of the Effective Date.

PARTNER 1:
Signature: _________________________
Name: [Printed Name]
Date: _____________________________

PARTNER 2:
Signature: _________________________
Name: [Printed Name]
Date: _____________________________

[Additional Partner Signatures]


EXHIBIT A: INITIAL VALUATION AND CAPITAL ACCOUNTS

[Details of initial contributions and valuations]

EXHIBIT B: BUSINESS PLAN

[Business plan, if applicable]
    `
  },
  {
    id: 'legal-mutual-nda',
    title: 'Confidentiality Agreement (Mutual NDA)',
    description: 'Mutual Non-Disclosure Agreement for protecting confidential information shared between two parties',
    category: 'legal',
    downloadCount: 12450,
    fileSize: '38.6 KB',
    rating: 4.9,
    tags: ['nda', 'confidentiality', 'legal', 'mutual', 'non-disclosure'],
    createdAt: '2024-01-10T09:15:00Z',
    updatedAt: '2024-01-31T14:00:00Z',
    content: `
MUTUAL NON-DISCLOSURE AGREEMENT

This Mutual Non-Disclosure Agreement ("Agreement") is entered into as of [Date] ("Effective Date") by and between:

PARTY 1:
[Company/Individual Name]
[Address]
[City, State, ZIP Code]
[Email]

PARTY 2:
[Company/Individual Name]
[Address]
[City, State, ZIP Code]
[Email]

(Each a "Party" and collectively the "Parties")

RECITALS

WHEREAS, the Parties wish to explore a business relationship concerning [purpose/project description] ("Purpose");

WHEREAS, in connection with the Purpose, each Party may disclose to the other certain confidential and proprietary information;

WHEREAS, the Parties desire to protect such confidential information from unauthorized use and disclosure;

NOW, THEREFORE, in consideration of the mutual covenants and agreements herein, the Parties agree:

1. DEFINITION OF CONFIDENTIAL INFORMATION

1.1 "Confidential Information" means any information disclosed by one Party (the "Disclosing Party") to the other Party (the "Receiving Party"), whether orally, in writing, electronically, or by visual observation, that:

(a) Is marked as "Confidential," "Proprietary," or with a similar designation; or

(b) Would reasonably be considered confidential given the nature of the information and circumstances of disclosure; or

(c) Includes, but is not limited to:
• Technical data, trade secrets, know-how, research, product plans, products, services, inventions, designs, processes, drawings, engineering, hardware, software, algorithms, source code
• Business information, pricing, costs, profits, markets, sales, customers, prospects, marketing plans, strategies
• Financial information, budgets, projections, investments
• Personnel information
• Any other proprietary information

1.2 Visual Observations
Information disclosed through facility tours, demonstrations, or observation of products, processes, or equipment shall be considered Confidential Information.

2. EXCLUSIONS FROM CONFIDENTIAL INFORMATION

Confidential Information does not include information that:

(a) Was known to the Receiving Party prior to disclosure by the Disclosing Party, as evidenced by written records;

(b) Is or becomes publicly available through no breach of this Agreement by the Receiving Party;

(c) Is rightfully received by the Receiving Party from a third party without breach of any confidentiality obligation;

(d) Is independently developed by the Receiving Party without use of or reference to the Disclosing Party's Confidential Information, as evidenced by written records;

(e) Is approved for release by written authorization of the Disclosing Party.

3. OBLIGATIONS OF RECEIVING PARTY

3.1 Non-Use
The Receiving Party shall use the Confidential Information solely for the Purpose and shall not use it for any other purpose without the prior written consent of the Disclosing Party.

3.2 Non-Disclosure
The Receiving Party shall:

(a) Maintain the Confidential Information in strict confidence using at least the same degree of care it uses to protect its own confidential information, but in no event less than reasonable care;

(b) Not disclose the Confidential Information to any third party except as permitted under Section 3.3;

(c) Protect the Confidential Information from unauthorized use, access, or disclosure.

3.3 Permitted Disclosures
The Receiving Party may disclose Confidential Information only to:

(a) Its employees, officers, directors, agents, contractors, and advisors (collectively "Representatives") who:
• Have a legitimate need to know for the Purpose;
• Have been informed of the confidential nature of the information; and
• Are bound by confidentiality obligations at least as restrictive as those in this Agreement.

(b) The Receiving Party is responsible for any breach of this Agreement by its Representatives.

3.4 Protective Measures
The Receiving Party shall:
• Limit access to Confidential Information on a need-to-know basis
• Implement reasonable security measures to prevent unauthorized disclosure
• Promptly notify the Disclosing Party of any unauthorized use or disclosure
• Cooperate with the Disclosing Party to remedy any breach

4. COMPELLED DISCLOSURE

4.1 If the Receiving Party is required by law, regulation, court order, or government authority to disclose any Confidential Information, it shall:

(a) Promptly notify the Disclosing Party in writing (unless legally prohibited) to allow the Disclosing Party to seek a protective order or other remedy;

(b) Provide reasonable assistance to the Disclosing Party in opposing or limiting the disclosure;

(c) Disclose only the minimum Confidential Information legally required;

(d) Request confidential treatment of any disclosed information.

4.2 The obligations of Section 3 shall apply to any Confidential Information not actually disclosed pursuant to legal compulsion.

5. OWNERSHIP AND NO LICENSE

5.1 Ownership
All Confidential Information remains the sole property of the Disclosing Party. No transfer of ownership or intellectual property rights is granted under this Agreement.

5.2 No License
Nothing in this Agreement grants any license, right, or interest in any patent, copyright, trademark, trade secret, or other intellectual property right, except the limited right to use Confidential Information for the Purpose.

5.3 No Obligation
Neither Party is obligated to:
• Disclose any Confidential Information
• Enter into any further agreement
• Continue any business relationship

6. RETURN OR DESTRUCTION OF CONFIDENTIAL INFORMATION

6.1 Upon written request by the Disclosing Party, or upon termination of this Agreement, the Receiving Party shall promptly:

(a) Return all tangible materials containing Confidential Information; and

(b) At the Disclosing Party's option, either:
• Destroy all copies of Confidential Information (including electronic copies); or
• Return all copies to the Disclosing Party.

6.2 The Receiving Party shall provide written certification of compliance with Section 6.1 upon request.

6.3 The Receiving Party may retain one archival copy of Confidential Information solely for legal compliance purposes, subject to continued confidentiality obligations.

6.4 Retention for Legal Compliance
The Receiving Party may retain Confidential Information to the extent required by law or bona fide document retention policy, provided it remains subject to confidentiality obligations.

7. TERM AND TERMINATION

7.1 Term
This Agreement shall commence on the Effective Date and continue for [#] years, unless earlier terminated.

7.2 Termination
Either Party may terminate this Agreement at any time upon [#] days written notice to the other Party.

7.3 Survival
The obligations under this Agreement shall survive termination for:
• Trade secrets: For as long as the information qualifies as a trade secret under applicable law
• Other Confidential Information: [#] years from the date of disclosure

8. NO WARRANTY

ALL CONFIDENTIAL INFORMATION IS PROVIDED "AS IS." THE DISCLOSING PARTY MAKES NO WARRANTIES, EXPRESS OR IMPLIED, REGARDING THE ACCURACY, COMPLETENESS, OR PERFORMANCE OF THE CONFIDENTIAL INFORMATION, INCLUDING ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.

9. REMEDIES

9.1 Irreparable Harm
The Parties acknowledge that breach of this Agreement may cause irreparable harm for which monetary damages are an inadequate remedy.

9.2 Injunctive Relief
In addition to all other remedies available at law or in equity, the Disclosing Party shall be entitled to seek injunctive relief to prevent or restrain any actual or threatened breach, without the necessity of posting a bond.

9.3 Other Remedies
The Parties' remedies under this Agreement are cumulative and not exclusive.

10. NO AGENCY OR PARTNERSHIP

This Agreement does not create any agency, partnership, joint venture, or employment relationship between the Parties. Neither Party has authority to bind the other.

11. GENERAL PROVISIONS

11.1 Entire Agreement
This Agreement constitutes the entire agreement regarding confidentiality and supersedes all prior agreements and understandings.

11.2 Amendments
Amendments must be in writing and signed by both Parties.

11.3 Waiver
No waiver of any provision shall be effective unless in writing. Waiver of one breach does not constitute waiver of subsequent breaches.

11.4 Severability
If any provision is held invalid or unenforceable, the remaining provisions shall continue in full force and effect.

11.5 Assignment
Neither Party may assign this Agreement without the prior written consent of the other Party, except to a successor in connection with a merger, acquisition, or sale of all or substantially all assets.

11.6 Governing Law
This Agreement shall be governed by the laws of [State/Country], without regard to conflicts of law principles.

11.7 Jurisdiction and Venue
Any dispute shall be resolved exclusively in the courts of [Jurisdiction], and the Parties consent to personal jurisdiction therein.

11.8 Export Control
The Receiving Party shall comply with all applicable export control laws and regulations.

11.9 Notices
All notices shall be in writing and delivered to the addresses above, or as updated in writing by a Party.

11.10 Counterparts
This Agreement may be executed in counterparts, each of which shall be deemed an original.

11.11 Electronic Signatures
This Agreement may be executed by electronic signature, which shall be deemed an original signature.

12. MUTUAL OBLIGATIONS

The Parties acknowledge that:
• Each Party may both disclose and receive Confidential Information
• The obligations under this Agreement apply equally to both Parties
• Each Party may independently possess or develop similar information

13. SPECIAL PROVISIONS

[Any additional provisions specific to the parties or purpose]


IN WITNESS WHEREOF, the Parties have executed this Agreement as of the Effective Date.

PARTY 1:

Signature: _________________________
Name: [Printed Name]
Title: [Title]
Date: _____________________________

PARTY 2:

Signature: _________________________
Name: [Printed Name]
Title: [Title]
Date: _____________________________
    `
  },
  {
    id: 'legal-terms-conditions',
    title: 'Terms & Conditions (T&C)',
    description: 'Comprehensive Terms and Conditions template for websites and online services',
    category: 'legal',
    downloadCount: 15620,
    fileSize: '61.4 KB',
    rating: 4.8,
    tags: ['terms', 'conditions', 'legal', 'website', 'ecommerce'],
    createdAt: '2024-01-11T13:45:00Z',
    updatedAt: '2024-02-01T09:30:00Z',
    content: `
TERMS AND CONDITIONS

Last Updated: [Date]

PLEASE READ THESE TERMS AND CONDITIONS CAREFULLY BEFORE USING THIS WEBSITE/SERVICE.

1. INTRODUCTION AND ACCEPTANCE

1.1 Agreement to Terms
These Terms and Conditions ("Terms," "Terms and Conditions") constitute a legally binding agreement between you ("User," "you," or "your") and [Company Name] ("Company," "we," "us," or "our") concerning your access to and use of the [Website Name] website located at [URL] (the "Website") and any related services, features, content, or applications (collectively, the "Services").

1.2 Acceptance
By accessing or using the Services, you agree to be bound by these Terms. If you do not agree to these Terms, you must not access or use the Services.

1.3 Changes to Terms
We reserve the right to modify these Terms at any time. We will notify users of material changes via [email/website notice/other method]. Your continued use after such modifications constitutes acceptance of the updated Terms. You are responsible for reviewing these Terms periodically.

1.4 Additional Terms
Certain features or services may be subject to additional terms and conditions, which will be presented to you in connection with such features or services.

2. ELIGIBILITY AND ACCOUNT REGISTRATION

2.1 Eligibility
You must be at least [13/16/18] years old to use the Services. By using the Services, you represent and warrant that you meet this age requirement.

2.2 Parental Consent
Users between [13-18] years old must have permission from a parent or legal guardian.

2.3 Account Registration
To access certain features, you must register for an account by providing:
• Full legal name
• Valid email address
• Password
• [Other required information]

2.4 Account Accuracy
You agree to provide accurate, current, and complete information and to update it as necessary to maintain accuracy.

2.5 Account Security
You are responsible for:
• Maintaining the confidentiality of your account credentials
• All activities that occur under your account
• Notifying us immediately of any unauthorized use
• Not sharing your account with others

2.6 Account Termination
We reserve the right to suspend or terminate your account at any time for:
• Violation of these Terms
• Fraudulent, abusive, or illegal activity
• Other reasons at our sole discretion

3. USE OF SERVICES

3.1 License Grant
Subject to these Terms, we grant you a limited, non-exclusive, non-transferable, revocable license to access and use the Services for personal [and/or commercial] purposes.

3.2 Restrictions
You agree NOT to:
• Violate any applicable laws or regulations
• Infringe on any third-party rights
• Upload or transmit viruses, malware, or other harmful code
• Attempt to gain unauthorized access to systems or accounts
• Interfere with or disrupt the Services
• Use automated means (bots, scrapers) without permission
• Reverse engineer, decompile, or disassemble any part of the Services
• Remove or modify any copyright, trademark, or proprietary notices
• Use the Services for any illegal or unauthorized purpose
• Impersonate any person or entity
• Collect user information without consent
• Spam, harass, or abuse other users

3.3 User Conduct
You agree to:
• Comply with all applicable laws
• Respect the rights of others
• Use the Services in good faith
• Not engage in any conduct that restricts or inhibits others' use

4. USER CONTENT

4.1 Definition
"User Content" means any content you submit, upload, post, or transmit through the Services, including text, images, videos, audio, and other materials.

4.2 Ownership
You retain ownership of your User Content. However, by submitting User Content, you grant us a worldwide, non-exclusive, royalty-free, transferable, sublicensable license to use, reproduce, modify, adapt, publish, translate, create derivative works from, distribute, and display such User Content in connection with the Services.

4.3 License Duration
The license granted in Section 4.2 continues even after you stop using the Services, except for User Content you delete (provided deletion is available).

4.4 Responsibility for User Content
You are solely responsible for your User Content and the consequences of posting it. You represent and warrant that:
• You own or have necessary rights to your User Content
• Your User Content does not violate these Terms or any law
• Your User Content does not infringe third-party rights

4.5 Prohibited Content
You may not post User Content that:
• Is illegal, harmful, threatening, abusive, harassing, defamatory, or invasive of privacy
• Infringes intellectual property or other proprietary rights
• Contains viruses or harmful code
• Is spam or unauthorized advertising
• Promotes discrimination, hate speech, or violence
• Exploits or harms minors
• Impersonates any person or entity
• [Other prohibited content]

4.6 Content Moderation
We reserve the right (but have no obligation) to:
• Monitor, review, or remove User Content
• Suspend or terminate accounts for violations
• Cooperate with law enforcement

4.7 No Endorsement
We do not endorse User Content and make no representations regarding its accuracy, reliability, or suitability.

5. INTELLECTUAL PROPERTY RIGHTS

5.1 Company IP
The Services and all content, features, and functionality (excluding User Content) are owned by the Company and are protected by copyright, trademark, patent, trade secret, and other intellectual property laws.

5.2 Trademarks
[Company Name], [Logo], and other marks are trademarks of the Company. You may not use these without our prior written permission.

5.3 Copyright Infringement
We respect intellectual property rights. If you believe content on the Services infringes your copyright, please contact our designated agent:

[Copyright Agent Name]
[Address]
[Email]
[Phone]

Your notice must include:
• Identification of the copyrighted work
• Identification of the infringing material and its location
• Your contact information
• A statement of good faith belief that use is unauthorized
• A statement under penalty of perjury that information is accurate
• Your physical or electronic signature

5.4 DMCA Policy
We will respond to valid DMCA notices and may terminate repeat infringers' accounts.

6. PURCHASES AND PAYMENTS

[If applicable to e-commerce or paid services]

6.1 Pricing
All prices are in [Currency] and are subject to change without notice. We strive to ensure accuracy but errors may occur.

6.2 Payment
You agree to pay all fees and applicable taxes for purchases made through the Services using authorized payment methods.

6.3 Payment Methods
We accept: [List payment methods - credit cards, PayPal, etc.]

6.4 Billing
• Subscription services are billed [monthly/annually] in advance
• One-time purchases are billed immediately
• You authorize us to charge your payment method

6.5 Refunds and Cancellations
• Refund policy: [Details]
• Cancellation policy: [Details]
• Cancellation process: [Steps]

6.6 Free Trials
• Duration: [# days]
• Auto-conversion to paid subscription unless cancelled
• Cancellation before trial ends: No charge

7. THIRD-PARTY LINKS AND SERVICES

7.1 Third-Party Links
The Services may contain links to third-party websites or services not controlled by us.

7.2 No Responsibility
We are not responsible for the content, privacy policies, or practices of third-party sites. You access them at your own risk.

7.3 No Endorsement
Inclusion of links does not imply endorsement.

8. DISCLAIMERS

8.1 AS-IS BASIS
THE SERVICES ARE PROVIDED "AS IS" AND "AS AVAILABLE" WITHOUT WARRANTIES OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE, AND NON-INFRINGEMENT.

8.2 NO GUARANTEES
WE DO NOT WARRANT THAT:
• The Services will be uninterrupted, secure, or error-free
• The results obtained will be accurate or reliable
• Any errors will be corrected
• The Services will meet your requirements

8.3 Use at Your Own Risk
Your use of the Services is at your sole risk. You are responsible for any damage to your computer system or loss of data.

9. LIMITATION OF LIABILITY

9.1 EXCLUSION OF DAMAGES
TO THE MAXIMUM EXTENT PERMITTED BY LAW, IN NO EVENT SHALL THE COMPANY, ITS OFFICERS, DIRECTORS, EMPLOYEES, OR AGENTS BE LIABLE FOR ANY INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, OR PUNITIVE DAMAGES, INCLUDING BUT NOT LIMITED TO LOSS OF PROFITS, DATA, USE, GOODWILL, OR OTHER INTANGIBLE LOSSES, RESULTING FROM:
• Your use or inability to use the Services
• Any unauthorized access to or use of our servers and/or any personal information stored therein
• Any interruption or cessation of transmission to or from the Services
• Any bugs, viruses, or harmful code transmitted through the Services
• Any errors or omissions in any content or for any loss or damage incurred as a result of use of any content posted, emailed, transmitted, or otherwise made available through the Services
• User Content or the defamatory, offensive, or illegal conduct of any third party

9.2 LIABILITY CAP
THE COMPANY'S TOTAL LIABILITY SHALL NOT EXCEED THE GREATER OF:
• The amount you paid to us in the 12 months prior to the claim; or
• $[Amount]

9.3 Exceptions
Some jurisdictions do not allow the exclusion of certain warranties or limitation of liability, so the above limitations may not apply to you.

10. INDEMNIFICATION

You agree to indemnify, defend, and hold harmless the Company, its affiliates, and their respective officers, directors, employees, and agents from and against any claims, liabilities, damages, losses, costs, expenses, or fees (including reasonable attorneys' fees) arising from:
• Your use of the Services
• Your violation of these Terms
• Your violation of any rights of another
• Your User Content

11. DISPUTE RESOLUTION

11.1 Governing Law
These Terms shall be governed by the laws of [State/Country], without regard to conflict of law provisions.

11.2 Informal Resolution
Before filing a claim, you agree to contact us and attempt to resolve the dispute informally by sending written notice to [Contact Information].

11.3 Arbitration Agreement
Any dispute arising from these Terms or the Services shall be resolved through binding arbitration under the rules of [AAA/JAMS] in [Location], except that either party may bring suit in court to enjoin infringement of intellectual property rights.

11.4 Class Action Waiver
YOU AGREE THAT DISPUTES MUST BE BROUGHT INDIVIDUALLY AND NOT AS A PLAINTIFF OR CLASS MEMBER IN ANY PURPORTED CLASS OR REPRESENTATIVE PROCEEDING.

11.5 Jury Trial Waiver
YOU WAIVE YOUR RIGHT TO A JURY TRIAL.

OR (Alternative to Arbitration):

11.3 Jurisdiction
You agree to submit to the personal and exclusive jurisdiction of the courts located in [Jurisdiction].

12. TERMINATION

12.1 By You
You may terminate your account at any time by [process].

12.2 By Us
We may suspend or terminate your access immediately, without prior notice, for any reason, including breach of these Terms.

12.3 Effect of Termination
Upon termination:
• Your right to use the Services ceases immediately
• We may delete your account and User Content
• Sections [#] survive termination

13. GENERAL PROVISIONS

13.1 Entire Agreement
These Terms constitute the entire agreement between you and the Company regarding the Services.

13.2 Waiver
Our failure to enforce any right or provision shall not constitute a waiver.

13.3 Severability
If any provision is held invalid or unenforceable, the remaining provisions continue in effect.

13.4 Assignment
You may not assign or transfer these Terms. We may assign our rights without restriction.

13.5 Force Majeure
We are not liable for delays or failures caused by circumstances beyond our reasonable control.

13.6 Notices
Notices to you may be sent to your email address on file. Notices to us should be sent to:

[Company Name]
[Address]
[Email]

13.7 Export Control
You agree to comply with all export and import laws and regulations.

13.8 U.S. Government Rights
If you are a U.S. government entity, the Services are "Commercial Items" as defined in FAR 2.101.

14. PRIVACY

Your use of the Services is also governed by our Privacy Policy, available at [URL]. Please review it to understand our data practices.

15. CONTACT INFORMATION

If you have questions about these Terms, please contact us:

[Company Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]
[Website]


By using the Services, you acknowledge that you have read, understood, and agree to be bound by these Terms and Conditions.
    `
  },
  {
    id: 'legal-privacy-policy',
    title: 'Privacy Policy (GDPR/CCPA-ready)',
    description: 'Comprehensive privacy policy template compliant with GDPR and CCPA regulations',
    category: 'legal',
    downloadCount: 14830,
    fileSize: '67.2 KB',
    rating: 4.9,
    tags: ['privacy', 'gdpr', 'ccpa', 'legal', 'data-protection'],
    createdAt: '2024-01-12T10:00:00Z',
    updatedAt: '2024-02-02T11:15:00Z',
    content: `
PRIVACY POLICY

Last Updated: [Date]
Effective Date: [Date]

INTRODUCTION

[Company Name] ("we," "us," or "our") is committed to protecting your privacy. This Privacy Policy explains how we collect, use, disclose, and safeguard your information when you visit our website [URL] and use our services (collectively, the "Services").

Please read this Privacy Policy carefully. By accessing or using the Services, you acknowledge that you have read, understood, and agree to be bound by this Privacy Policy. If you do not agree, please do not access or use the Services.

TABLE OF CONTENTS

1. Information We Collect
2. How We Use Your Information
3. How We Share Your Information
4. Cookies and Tracking Technologies
5. Third-Party Services
6. Data Security
7. Data Retention
8. Your Privacy Rights
9. Children's Privacy
10. International Data Transfers
11. California Privacy Rights (CCPA)
12. European Privacy Rights (GDPR)
13. Do Not Track Signals
14. Changes to This Privacy Policy
15. Contact Us

1. INFORMATION WE COLLECT

We collect information that you provide directly, information collected automatically, and information from third parties.

1.1 Information You Provide Directly

A. Account Information
When you create an account, we collect:
• Name
• Email address
• Password (encrypted)
• Phone number
• [Other account details]

B. Profile Information
• Profile photo
• Bio or description
• Preferences and settings
• [Other profile details]

C. Payment Information
• Billing address
• Payment card information (processed by [Payment Processor])
• Transaction history

D. Communications
• Customer support inquiries
• Survey responses
• Feedback and reviews
• Communications you send us

E. User Content
• Posts, comments, messages
• Files, photos, videos you upload
• [Other content]

1.2 Information Collected Automatically

A. Device and Usage Information
• IP address
• Browser type and version
• Operating system
• Device identifiers
• Pages visited and time spent
• Referring URLs
• Click patterns and interactions

B. Location Information
• Approximate location based on IP address
• Precise location (if you grant permission)

C. Cookies and Similar Technologies
We use cookies, web beacons, pixel tags, and similar technologies to collect information. See Section 4 for details.

1.3 Information from Third Parties

A. Social Media
If you connect via social media (Facebook, Google, etc.), we may receive:
• Profile information
• Email address
• Friends/connections list
• [Other information based on your privacy settings]

B. Third-Party Services
We may receive information from:
• Analytics providers
• Advertising partners
• Data brokers (where permitted by law)
• Publicly available sources

2. HOW WE USE YOUR INFORMATION

We use your information for the following purposes:

2.1 Provide and Improve Services
• Create and manage your account
• Process transactions
• Provide customer support
• Personalize your experience
• Improve and optimize the Services
• Develop new features and services

2.2 Communications
• Send service-related announcements
• Respond to inquiries
• Send marketing communications (with your consent)
• Send newsletters (with your consent)
• Conduct surveys

2.3 Safety and Security
• Detect and prevent fraud, abuse, and illegal activity
• Enforce our Terms of Service
• Protect our rights and property
• Comply with legal obligations
• Ensure safety of users

2.4 Analytics and Research
• Analyze usage patterns and trends
• Conduct research and analysis
• Create aggregated, de-identified data

2.5 Advertising and Marketing
• Display targeted advertisements
• Measure ad effectiveness
• Send promotional offers (with your consent)

2.6 Legal Compliance
• Comply with legal obligations
• Respond to legal requests
• Enforce our policies
• Protect rights, property, and safety

3. HOW WE SHARE YOUR INFORMATION

We share your information in the following circumstances:

3.1 With Your Consent
We share information when you direct us to do so.

3.2 Service Providers
We share information with third-party service providers who perform services on our behalf:
• Payment processors
• Cloud hosting providers
• Email service providers
• Analytics providers
• Customer support tools
• Marketing platforms

These providers are contractually obligated to use your information only for the purposes of providing services to us and to protect your information.

3.3 Business Transfers
If we are involved in a merger, acquisition, sale of assets, or bankruptcy, your information may be transferred as part of that transaction.

3.4 Legal Requirements
We may disclose information to:
• Comply with laws, regulations, or legal processes
• Respond to government requests
• Enforce our Terms of Service
• Protect our rights, property, or safety
• Prevent fraud or illegal activity

3.5 Aggregated or De-Identified Data
We may share aggregated or de-identified information that cannot reasonably be used to identify you.

3.6 Public Information
Information you post publicly (e.g., comments, reviews) may be viewed by others.

3.7 With Affiliates
We may share information with our parent company, subsidiaries, and affiliates.

4. COOKIES AND TRACKING TECHNOLOGIES

4.1 What Are Cookies?
Cookies are small text files stored on your device that help us recognize you and remember your preferences.

4.2 Types of Cookies We Use

A. Essential Cookies
Required for the Services to function. You cannot opt out of these.

B. Functional Cookies
Remember your preferences and settings.

C. Analytics Cookies
Help us understand how you use the Services (e.g., Google Analytics).

D. Advertising Cookies
Used to deliver targeted ads based on your interests.

4.3 Cookie Management
You can control cookies through your browser settings. Note that disabling cookies may affect functionality.

4.4 Other Tracking Technologies
We use web beacons, pixel tags, and similar technologies for analytics and advertising.

4.5 Do Not Track
See Section 13 for our Do Not Track policy.

5. THIRD-PARTY SERVICES

5.1 Third-Party Links
The Services may contain links to third-party websites and services. We are not responsible for their privacy practices.

5.2 Third-Party Plugins
We use social media plugins (e.g., Facebook Like button) that may collect information about you.

5.3 Analytics Services
We use:
• Google Analytics [Opt-out: https://tools.google.com/dlpage/gaoptout]
• [Other analytics services]

5.4 Advertising Networks
We work with advertising partners who may use cookies and similar technologies. You can opt out at:
• Network Advertising Initiative: http://www.networkadvertising.org/choices/
• Digital Advertising Alliance: http://www.aboutads.info/choices/
• European Interactive Digital Advertising Alliance: http://www.youronlinechoices.eu/

6. DATA SECURITY

6.1 Security Measures
We implement reasonable administrative, technical, and physical security measures to protect your information, including:
• Encryption of data in transit (SSL/TLS)
• Encryption of sensitive data at rest
• Regular security assessments
• Access controls and authentication
• Employee training

6.2 No Guarantee
However, no method of transmission or storage is 100% secure. We cannot guarantee absolute security.

6.3 Your Responsibility
You are responsible for maintaining the confidentiality of your account credentials.

6.4 Breach Notification
In the event of a data breach, we will notify affected users as required by law.

7. DATA RETENTION

7.1 Retention Period
We retain your information for as long as necessary to:
• Provide the Services
• Comply with legal obligations
• Resolve disputes
• Enforce our policies

7.2 Deletion
When we no longer need your information, we will securely delete or anonymize it.

7.3 Backup Copies
Some information may remain in backup copies for a limited time.

8. YOUR PRIVACY RIGHTS

Depending on your location, you may have the following rights:

8.1 Access
You may request a copy of your personal information.

8.2 Correction
You may request that we correct inaccurate information.

8.3 Deletion
You may request that we delete your information, subject to certain exceptions.

8.4 Portability
You may request your information in a portable format.

8.5 Objection
You may object to certain processing of your information.

8.6 Restriction
You may request that we restrict certain processing.

8.7 Withdraw Consent
If processing is based on consent, you may withdraw it at any time.

8.8 Opt Out of Marketing
You may opt out of marketing communications by:
• Clicking "unsubscribe" in emails
• Updating your account settings
• Contacting us

8.9 How to Exercise Your Rights
To exercise these rights, contact us at:
[Privacy Contact Email]
[Privacy Contact Address]

We will respond within [30] days.

9. CHILDREN'S PRIVACY

9.1 Age Restriction
The Services are not intended for children under [13/16] years of age. We do not knowingly collect information from children.

9.2 Parental Notice
If you are a parent and believe your child has provided us with information, please contact us.

9.3 Deletion
If we learn we have collected information from a child without parental consent, we will delete it promptly.

10. INTERNATIONAL DATA TRANSFERS

10.1 Global Operations
We operate globally and may transfer your information to countries other than your own, including [Country where servers are located].

10.2 Adequate Protection
We ensure adequate protection through:
• EU-U.S. Privacy Shield (if applicable)
• Standard Contractual Clauses
• Adequacy decisions
• Your consent

10.3 European Users
For European users, data may be transferred outside the EEA. We use appropriate safeguards as required by GDPR.

11. CALIFORNIA PRIVACY RIGHTS (CCPA)

If you are a California resident, you have additional rights under the California Consumer Privacy Act (CCPA):

11.1 Right to Know
You have the right to request:
• Categories of personal information collected
• Categories of sources
• Business purpose for collection
• Categories of third parties with whom we share information
• Specific pieces of personal information

11.2 Right to Delete
You have the right to request deletion of your personal information, subject to exceptions.

11.3 Right to Opt-Out of Sale
We [do/do not] sell personal information as defined by CCPA. If we do, you have the right to opt out at: [Do Not Sell My Personal Information Link]

11.4 Right to Non-Discrimination
You have the right not to receive discriminatory treatment for exercising your CCPA rights.

11.5 Authorized Agent
You may designate an authorized agent to make requests on your behalf.

11.6 Verification
We will verify your identity before processing requests.

11.7 California Shine the Light
California residents may request information about our disclosure of personal information to third parties for their direct marketing purposes. Contact us at [Email].

12. EUROPEAN PRIVACY RIGHTS (GDPR)

If you are in the European Economic Area (EEA), you have rights under the General Data Protection Regulation (GDPR):

12.1 Legal Basis for Processing
We process your information based on:
• Consent
• Performance of a contract
• Legitimate interests
• Legal obligations

12.2 Data Controller
[Company Name] is the data controller for your personal data.

Data Protection Officer Contact:
[DPO Name]
[DPO Email]
[DPO Address]

12.3 Your GDPR Rights
• Right of access
• Right to rectification
• Right to erasure ("right to be forgotten")
• Right to restriction of processing
• Right to data portability
• Right to object
• Rights related to automated decision-making and profiling

12.4 Supervisory Authority
You have the right to lodge a complaint with your local data protection authority.

12.5 Cross-Border Data Transfers
See Section 10.

13. DO NOT TRACK SIGNALS

13.1 Browser DNT
Some browsers have a "Do Not Track" feature. We [do/do not] currently respond to DNT signals.

13.2 Global Privacy Control
We [do/do not] honor Global Privacy Control (GPC) signals.

14. CHANGES TO THIS PRIVACY POLICY

14.1 Updates
We may update this Privacy Policy from time to time. We will notify you of material changes by:
• Posting the updated Privacy Policy with a new "Last Updated" date
• Sending you an email notification
• Displaying a prominent notice on the Services

14.2 Acceptance
Your continued use of the Services after changes constitutes acceptance of the updated Privacy Policy.

14.3 Review
We encourage you to review this Privacy Policy periodically.

15. CONTACT US

If you have questions, concerns, or requests regarding this Privacy Policy or our privacy practices, please contact us:

[Company Name]
Attn: Privacy Department
[Address]
[City, State, ZIP Code]

Email: [Privacy Email]
Phone: [Privacy Phone]
Website: [Privacy Webpage]

For GDPR-related inquiries, contact our Data Protection Officer:
[DPO Name]
[DPO Email]

For CCPA-related inquiries:
[CCPA Contact Email]


ADDITIONAL DISCLOSURES FOR SPECIFIC JURISDICTIONS

[Add any additional disclosures required by specific states or countries]


By using our Services, you acknowledge that you have read and understood this Privacy Policy.
    `
  },
  {
    id: 'legal-ip-assignment',
    title: 'Intellectual Property Assignment Agreement',
    description: 'Agreement for transferring intellectual property rights from creator to company or purchaser',
    category: 'legal',
    downloadCount: 6890,
    fileSize: '44.5 KB',
    rating: 4.7,
    tags: ['ip', 'intellectual-property', 'assignment', 'legal', 'copyright'],
    createdAt: '2024-01-13T14:20:00Z',
    updatedAt: '2024-02-03T16:00:00Z',
    content: `
INTELLECTUAL PROPERTY ASSIGNMENT AGREEMENT

This Intellectual Property Assignment Agreement ("Agreement") is entered into as of [Date] ("Effective Date") by and between:

ASSIGNOR:
[Full Legal Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]

ASSIGNEE:
[Company/Individual Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]

RECITALS

WHEREAS, Assignor has created, developed, or owns certain intellectual property;

WHEREAS, Assignee desires to acquire all rights, title, and interest in such intellectual property;

WHEREAS, Assignor desires to assign such rights to Assignee in exchange for the consideration set forth herein;

NOW, THEREFORE, in consideration of the mutual covenants and agreements herein, the parties agree:

1. DEFINITIONS

1.1 "Assigned IP" means all intellectual property described in Exhibit A and all intellectual property created, developed, or conceived by Assignor:
☐ During the term of employment/engagement with Assignee
☐ Related to the Project (as defined below)
☐ [Other scope]

1.2 "Intellectual Property" or "IP" includes:
• Inventions, discoveries, and improvements (whether patentable or not)
• Works of authorship (copyrightable works)
• Trade secrets and confidential information
• Trademarks, service marks, trade names, and logos
• Designs, processes, methods, and know-how
• Software, code, algorithms, and documentation
• Databases and compilations
• All rights in the foregoing worldwide

1.3 "Project" means: [Description of project, if applicable]

2. ASSIGNMENT OF INTELLECTUAL PROPERTY

2.1 Assignment
Assignor hereby irrevocably assigns, transfers, and conveys to Assignee all right, title, and interest in and to the Assigned IP, including:

A. Ownership Rights
• All ownership rights worldwide
• All rights to use, reproduce, modify, distribute, display, and perform
• All rights to create derivative works
• All rights to sublicense and transfer

B. Intellectual Property Rights
• All patent rights and applications
• All copyright and related rights
• All trademark rights
• All trade secret rights
• All moral rights (to the extent assignable)
• All other proprietary rights

C. Embodiments
• All tangible embodiments (documents, prototypes, code, etc.)
• All electronic files and data
• All notes, drawings, and specifications
• All related materials

2.2 Prior Works
This Agreement does not apply to the following pre-existing intellectual property owned by Assignor:
[List any excluded prior works, or state "None"]

2.3 License to Prior Works
☐ N/A - No prior works incorporated
☐ Assignor grants Assignee a [exclusive/non-exclusive], [perpetual/term-limited], [royalty-free/royalty-bearing], worldwide license to use Prior Works to the extent incorporated in Assigned IP for [permitted uses].

3. CONSIDERATION

In consideration for this assignment, Assignee shall provide:

3.1 Payment
☐ One-time payment of $[Amount], payable [upon execution/within # days]
☐ Royalty of [%] of [net revenue/gross sales] from Assigned IP, paid [quarterly/annually]
☐ Employment compensation (IP created during employment)
☐ Other: [Specify consideration]

3.2 Additional Consideration
☐ Stock/equity: [# shares or %]
☐ Attribution/credit: [Details]
☐ Other benefits: [Specify]

3.3 Payment Terms
• Method: [Wire, check, etc.]
• Schedule: [Details]
• Conditions: [If applicable]

4. REPRESENTATIONS AND WARRANTIES

4.1 Assignor's Representations
Assignor represents and warrants that:

A. Ownership
• Assignor is the sole and exclusive owner of the Assigned IP
• Assignor has full right and authority to assign the Assigned IP
• No other person or entity has any rights in the Assigned IP

B. Originality
• The Assigned IP is original to Assignor
• The Assigned IP does not infringe any third-party rights
• Assignor has not licensed or assigned the Assigned IP to any third party

C. No Conflicts
• This assignment does not violate any agreement with third parties
• Assignor is not subject to any conflicting obligations
• No consents from third parties are required

D. No Litigation
• There are no pending or threatened claims regarding the Assigned IP
• Assignor is not aware of any infringement by third parties

E. Quality
• [If applicable] The Assigned IP is free from material defects
• [If applicable] The Assigned IP performs substantially as described

4.2 DISCLAIMER
EXCEPT AS EXPRESSLY PROVIDED IN SECTION 4.1, THE ASSIGNED IP IS PROVIDED "AS IS" WITHOUT ANY WARRANTY, EXPRESS OR IMPLIED, INCLUDING WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.

4.3 Assignee's Representations
Assignee represents that it has authority to enter into this Agreement.

5. ASSIGNOR'S OBLIGATIONS

5.1 Delivery of Materials
Assignor shall promptly deliver to Assignee:
• All tangible embodiments of the Assigned IP
• All documentation, specifications, and notes
• All source code, files, and data
• All passwords, access credentials, and keys
• Any other materials related to the Assigned IP

5.2 Further Assurances
Assignor agrees to:
• Execute any documents reasonably necessary to perfect Assignee's rights
• Cooperate in obtaining patents, copyrights, or other registrations
• Assist in enforcement against infringers
• Provide testimony or declarations if needed
• Take any other actions reasonably requested by Assignee

5.3 Cooperation
Assignor shall cooperate at Assignee's expense in:
• Patent prosecution
• Copyright registration
• Litigation or disputes
• Due diligence requests

5.4 Moral Rights Waiver
To the extent permitted by law, Assignor irrevocably waives all moral rights (including attribution, integrity, and withdrawal rights) in the Assigned IP.

If moral rights cannot be waived, Assignor agrees not to assert them against Assignee or its successors.

6. ASSIGNEE'S RIGHTS

6.1 Unrestricted Use
Assignee shall have the unrestricted right to:
• Use, modify, and create derivative works
• Commercialize and monetize the Assigned IP
• License or assign to third parties
• Register patents, copyrights, or trademarks
• Enforce rights against infringers
• Take any other action an owner may take

6.2 No Obligation
Assignee has no obligation to:
• Use or commercialize the Assigned IP
• Provide attribution to Assignor
• Continue development
• Pay royalties beyond those specified in Section 3

7. CONFIDENTIALITY

7.1 Confidential Information
Assignor agrees to maintain in confidence all proprietary and confidential information of Assignee disclosed in connection with this Agreement.

7.2 Obligations
Assignor shall:
• Use confidential information only for purposes of this Agreement
• Protect it with reasonable care
• Not disclose to third parties without consent

7.3 Exclusions
Confidentiality does not apply to information that:
• Is publicly available
• Was known prior to disclosure
• Is independently developed
• Must be disclosed by law

7.4 Duration
Confidentiality obligations survive for [#] years after termination.

8. INDEMNIFICATION

8.1 Assignor's Indemnification
Assignor shall indemnify, defend, and hold harmless Assignee from claims arising from:
• Breach of Assignor's representations and warranties
• Infringement of third-party rights by the Assigned IP
• Assignor's negligence or willful misconduct

8.2 Assignee's Indemnification
Assignee shall indemnify Assignor from claims arising from:
• Assignee's use of the Assigned IP after assignment
• Modifications made by Assignee
• [Other Assignee liabilities]

8.3 Indemnification Procedure
Indemnified party must:
• Promptly notify indemnifying party
• Allow indemnifying party to control defense
• Cooperate reasonably

9. LIMITATION OF LIABILITY

9.1 ASSIGNOR'S LIABILITY CAP
ASSIGNOR'S TOTAL LIABILITY SHALL NOT EXCEED:
☐ The consideration paid under Section 3
☐ $[Amount]
☐ [Other limitation]

9.2 EXCLUDED DAMAGES
NEITHER PARTY SHALL BE LIABLE FOR INDIRECT, INCIDENTAL, CONSEQUENTIAL, OR PUNITIVE DAMAGES, INCLUDING LOST PROFITS.

9.3 Exceptions
Limitations do not apply to:
• Willful misconduct or fraud
• Breach of confidentiality
• Indemnification obligations
• [Other exceptions]

10. TERM AND TERMINATION

10.1 Term
This Agreement is effective as of the Effective Date and continues perpetually, except for:
• Assignor's obligations under Section 5: [Term or perpetual]
• Confidentiality obligations: [# years]

10.2 Termination
☐ This Agreement cannot be terminated
☐ Assignee may terminate if Assignor materially breaches
☐ Other termination rights: [Specify]

10.3 Effect of Termination
☐ N/A - Agreement is perpetual
☐ Upon termination, assignment remains effective but [other effects]

11. GENERAL PROVISIONS

11.1 Entire Agreement
This Agreement constitutes the entire agreement and supersedes all prior agreements.

11.2 Amendments
Amendments must be in writing and signed by both parties.

11.3 Assignment
• Assignor may not assign this Agreement without consent
• Assignee may freely assign this Agreement
• This Agreement binds successors and assigns

11.4 Governing Law
This Agreement shall be governed by the laws of [State/Country].

11.5 Dispute Resolution
☐ Mediation required before litigation
☐ Binding arbitration under [AAA/JAMS] rules in [Location]
☐ Litigation in courts of [Jurisdiction]

11.6 Attorney's Fees
Prevailing party in dispute entitled to reasonable attorney's fees.

11.7 Severability
Invalid provisions severed; remainder continues in effect.

11.8 Waiver
Waiver must be in writing and does not constitute ongoing waiver.

11.9 Notices
Notices shall be sent to addresses above or as updated in writing.

11.10 Counterparts
This Agreement may be executed in counterparts and by electronic signature.

11.11 Survival
Sections [#] survive termination or expiration of this Agreement.

12. SPECIAL PROVISIONS

[Any additional provisions specific to the assignment]


IN WITNESS WHEREOF, the parties have executed this Agreement as of the Effective Date.

ASSIGNOR:

Signature: _________________________
Name: [Printed Name]
Date: _____________________________

ASSIGNEE:

Signature: _________________________
Name: [Printed Name]
Title: [Title]
Date: _____________________________


EXHIBIT A: DESCRIPTION OF ASSIGNED IP

[Detailed description of the intellectual property being assigned, including:
- Inventions and patents (with patent numbers or applications)
- Copyrighted works (with registration numbers if applicable)
- Trademarks (with registration numbers if applicable)
- Trade secrets and know-how
- Software and code (with repositories, versions)
- Designs and drawings
- Any other IP being assigned]


EXHIBIT B: PRIOR WORKS (IF APPLICABLE)

[List of Assignor's prior works that are excluded from this assignment]


EXHIBIT C: PAYMENT SCHEDULE (IF APPLICABLE)

[Detailed payment terms, milestones, and amounts]
    `
  },
  {
    id: 'legal-independent-contractor',
    title: 'Independent Contractor Agreement',
    description: 'Comprehensive agreement for engaging independent contractors with clear terms and deliverables',
    category: 'legal',
    downloadCount: 11240,
    fileSize: '55.8 KB',
    rating: 4.8,
    tags: ['contractor', 'freelance', 'legal', 'agreement', '1099'],
    createdAt: '2024-01-14T11:00:00Z',
    updatedAt: '2024-02-04T13:45:00Z',
    content: `
INDEPENDENT CONTRACTOR AGREEMENT

This Independent Contractor Agreement ("Agreement") is entered into as of [Date] ("Effective Date") by and between:

COMPANY:
[Company Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]

CONTRACTOR:
[Full Legal Name / Business Name]
[Address]
[City, State, ZIP Code]
[Email]
[Phone]
Tax ID/SSN: [Tax ID Number]

RECITALS

WHEREAS, Company desires to engage Contractor to perform certain services;
WHEREAS, Contractor is in the business of providing such services;
WHEREAS, the parties intend to establish an independent contractor relationship;

NOW, THEREFORE, in consideration of the mutual covenants herein, the parties agree:

1. SERVICES

1.1 Scope of Services
Contractor agrees to provide the following services ("Services"):

[Detailed description of services, deliverables, and specifications]

Examples:
• [Service/deliverable 1]
• [Service/deliverable 2]
• [Service/deliverable 3]
• [Additional services as specified in Exhibit A]

1.2 Performance Standards
Services shall be performed:
• In a professional and workmanlike manner
• In accordance with industry standards
• According to specifications in Exhibit A
• In compliance with all applicable laws

1.3 Work Product
"Work Product" means all deliverables, documents, materials, and intellectual property created by Contractor in performing the Services.

2. TERM AND TERMINATION

2.1 Term
This Agreement shall commence on [Start Date] and continue until:
☐ [End Date]
☐ Completion of Services
☐ Termination by either party

2.2 Termination for Convenience
Either party may terminate this Agreement with [#] days written notice.

2.3 Termination for Cause
Either party may terminate immediately upon written notice if the other party:
• Materially breaches this Agreement and fails to cure within [#] days
• Becomes insolvent or files for bankruptcy
• Engages in fraud, illegal activities, or gross negligence

2.4 Effect of Termination
Upon termination:
• Contractor shall deliver all completed Work Product
• Company shall pay for services satisfactorily performed through termination date
• Contractor shall return all Company property and confidential information
• Sections [3, 4, 5, 6, 7, 8] survive termination

3. COMPENSATION AND PAYMENT

3.1 Fees
Company agrees to pay Contractor as follows:

☐ Fixed Fee: $[Amount] for [scope of work]
  Payable: ☐ Upon completion ☐ In installments per Exhibit B

☐ Hourly Rate: $[Rate] per hour
  Estimated hours: [#]
  Maximum without approval: [#] hours or $[Amount]

☐ Per Project/Milestone: As specified in Exhibit B

☐ Monthly Retainer: $[Amount] per month for [scope]

☐ Other: [Specify payment structure]

3.2 Expenses
☐ Expenses are included in the fees
☐ Pre-approved expenses reimbursed (estimated: $[Amount])
  • Reimbursement requires receipts
  • Maximum reimbursable: $[Amount] without prior approval
☐ Contractor pays all expenses

3.3 Invoicing
• Frequency: [Weekly/Monthly/Upon milestones/Upon completion]
• Invoice must include: [Details required - hours, description, receipts]
• Submit to: [Email/address]

3.4 Payment Terms
• Payment due: Net [#] days from invoice date
• Payment method: [ACH, check, wire transfer, etc.]
• Late payment: Interest of [%] per month on overdue amounts

3.5 Taxes
Contractor is responsible for all taxes, including:
• Income taxes
• Self-employment taxes
• Sales/use taxes
• Other applicable taxes

Company will issue Form 1099-NEC if payments exceed $600 in a calendar year.

4. INDEPENDENT CONTRACTOR RELATIONSHIP

4.1 Independent Contractor Status
Contractor is an independent contractor, not an employee, partner, or agent of Company.

4.2 No Employment Benefits
Contractor is not entitled to:
• Health insurance, retirement, or other employee benefits
• Unemployment insurance
• Workers' compensation
• Paid time off
• Any other employee benefits

4.3 No Authority to Bind
Contractor has no authority to bind Company or enter into agreements on Company's behalf.

4.4 Contractor's Business
Contractor:
• Maintains their own business
• May provide services to other clients (subject to Section 6)
• Controls the manner and means of performing Services (subject to deliverable specifications)
• Provides their own tools, equipment, and workspace

4.5 IRS Factors
The parties intend this to be an independent contractor relationship under IRS guidelines and acknowledge:
• Contractor controls how Services are performed
• Contractor is engaged in an independently established business
• Contractor provides similar services to others
• Contractor has the opportunity for profit or loss
• Contractor invests in their own business
• The relationship is not indefinite

5. INTELLECTUAL PROPERTY

5.1 Ownership of Work Product
☐ Work-Made-for-Hire: All Work Product is a "work made for hire" and Company owns all rights.

☐ Assignment: Contractor assigns all right, title, and interest in Work Product to Company upon [full payment/creation].

☐ License: Contractor retains ownership and grants Company a [exclusive/non-exclusive], [perpetual/term-limited], [royalty-free/royalty-bearing], worldwide license to use Work Product for [purposes].

5.2 Pre-Existing Materials
Contractor's pre-existing materials ("Pre-Existing IP") incorporated into Work Product:
[List or state "None"]

License to Pre-Existing IP:
Contractor grants Company a [exclusive/non-exclusive], [perpetual/term-limited], royalty-free license to use Pre-Existing IP to the extent incorporated in Work Product.

5.3 Contractor's Retained Rights
Contractor retains rights in:
• General knowledge, skills, and experience
• Methods, processes, and techniques of general application
• [Other retained rights]

5.4 Third-Party Materials
Contractor shall not incorporate third-party materials without Company's prior written approval.

5.5 Moral Rights
Contractor waives all moral rights in the Work Product to the extent permitted by law.

5.6 Further Assurances
Contractor agrees to execute any documents necessary to:
• Perfect Company's rights in Work Product
• Assist in obtaining patents, copyrights, or other registrations
• Enforce rights against infringers

6. CONFIDENTIALITY AND NON-DISCLOSURE

6.1 Confidential Information
"Confidential Information" means all proprietary information of Company, including:
• Business strategies, financial information, and customer data
• Technical data, trade secrets, and know-how
• Product plans and marketing strategies
• Information marked "Confidential" or reasonably understood to be confidential

6.2 Obligations
Contractor shall:
• Use Confidential Information only to perform Services
• Protect it with at least reasonable care
• Not disclose to third parties without written consent
• Limit access to those with a need to know

6.3 Exceptions
Obligations do not apply to information that:
• Is publicly available through no fault of Contractor
• Was rightfully known prior to disclosure
• Is independently developed without use of Confidential Information
• Must be disclosed by law (with prompt notice to Company)

6.4 Return of Information
Upon termination, Contractor shall return or destroy all Confidential Information and provide written certification.

6.5 Duration
Confidentiality obligations survive for [#] years after termination, or indefinitely for trade secrets.

7. NON-COMPETE AND NON-SOLICITATION

[Note: Enforceability varies by jurisdiction]

7.1 Non-Compete
During the term and for [#] months/years after termination, Contractor shall not:
☐ N/A - No non-compete restriction
☐ Directly compete with Company in [geographic area] by [specific restricted activities]

7.2 Non-Solicitation of Customers
During the term and for [#] months after termination, Contractor shall not solicit Company's customers with whom Contractor had contact or about whom Contractor learned confidential information.

☐ N/A - No customer non-solicitation

7.3 Non-Solicitation of Employees
During the term and for [#] months after termination, Contractor shall not solicit or hire Company's employees.

☐ N/A - No employee non-solicitation

7.4 Exception
These restrictions do not apply to:
☐ Services unrelated to those provided to Company
☐ General advertising not targeted at Company's customers/employees
☐ [Other exceptions]

8. REPRESENTATIONS AND WARRANTIES

8.1 Contractor's Warranties
Contractor represents and warrants that:
• Contractor has the right and ability to enter into this Agreement
• Services will be performed in a professional manner
• Work Product will be original and will not infringe third-party rights
• Contractor has appropriate licenses, permits, and insurance
• Contractor will comply with all applicable laws
• Contractor is not subject to any conflicting obligations

8.2 No Conflicts
Contractor represents that this Agreement does not violate any agreement with third parties or any legal restrictions.

8.3 Company's Warranties
Company represents that it has authority to enter this Agreement and will make timely payments.

8.4 DISCLAIMER
EXCEPT AS EXPRESSLY PROVIDED, SERVICES ARE PROVIDED "AS IS" WITHOUT WARRANTY. CONTRACTOR DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.

9. INDEMNIFICATION

9.1 Contractor's Indemnification
Contractor shall indemnify and defend Company against claims arising from:
• Contractor's negligence or willful misconduct
• Infringement of third-party IP rights by Work Product
• Breach of Contractor's representations and warranties
• Violation of laws by Contractor
• Injury or damage caused by Contractor

9.2 Company's Indemnification
Company shall indemnify Contractor against claims arising from:
• Company's negligence or willful misconduct
• Company's use of Work Product beyond scope of license
• Breach of Company's obligations

9.3 Procedure
Indemnified party must:
• Promptly notify indemnifying party of claim
• Allow indemnifying party to control defense
• Cooperate reasonably in defense

10. LIMITATION OF LIABILITY

10.1 LIABILITY CAP
CONTRACTOR'S TOTAL LIABILITY SHALL NOT EXCEED:
☐ Fees paid in the 12 months preceding the claim
☐ $[Amount]
☐ [Other limitation]

10.2 CONSEQUENTIAL DAMAGES
NEITHER PARTY SHALL BE LIABLE FOR INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, OR PUNITIVE DAMAGES, INCLUDING LOST PROFITS, EVEN IF ADVISED OF THEIR POSSIBILITY.

10.3 Exceptions
Limitations do not apply to:
• Willful misconduct or fraud
• Breach of confidentiality
• Indemnification obligations
• [Other exceptions]

11. INSURANCE

Contractor shall maintain the following insurance during the term:
☐ General Liability: $[Amount] per occurrence
☐ Professional Liability (E&O): $[Amount] per claim
☐ Workers' Compensation: As required by law (if Contractor has employees)
☐ Other: [Specify]

Contractor shall provide certificates of insurance upon request.

☐ N/A - No insurance required

12. GENERAL PROVISIONS

12.1 Entire Agreement
This Agreement constitutes the entire agreement and supersedes all prior agreements.

12.2 Amendments
Amendments must be in writing and signed by both parties.

12.3 Assignment
Contractor may not assign this Agreement without Company's written consent. Company may assign to a successor in merger or acquisition.

12.4 Independent Legal Advice
Each party has had opportunity to obtain independent legal and tax advice regarding this Agreement.

12.5 Governing Law
This Agreement shall be governed by the laws of [State/Country].

12.6 Dispute Resolution
☐ Mediation required before litigation in [Location]
☐ Binding arbitration under [AAA/JAMS] rules in [Location]
☐ Litigation in courts of [Jurisdiction]

12.7 Attorney's Fees
Prevailing party in dispute entitled to reasonable attorney's fees and costs.

12.8 Severability
Invalid provisions severed; remainder continues in effect.

12.9 Waiver
Failure to enforce any provision does not waive right to enforce later.

12.10 Notices
Notices shall be sent to addresses above or as updated in writing.

12.11 Force Majeure
Neither party liable for delays caused by events beyond reasonable control.

12.12 Counterparts
This Agreement may be executed in counterparts and by electronic signature.


IN WITNESS WHEREOF, the parties have executed this Agreement as of the Effective Date.

COMPANY:

Signature: _________________________
Name: [Printed Name]
Title: [Title]
Date: _____________________________

CONTRACTOR:

Signature: _________________________
Name: [Printed Name]
Title/DBA: [If applicable]
Date: _____________________________


EXHIBIT A: STATEMENT OF WORK

[Detailed description of Services, deliverables, specifications, milestones, acceptance criteria, timelines, etc.]


EXHIBIT B: PAYMENT SCHEDULE

[Detailed payment terms, milestones, amounts, and due dates]


EXHIBIT C: PRE-EXISTING INTELLECTUAL PROPERTY

[List of Contractor's pre-existing materials, if any]
    `
  },
  {
    id: 'legal-power-of-attorney',
    title: 'Power of Attorney (Limited Use)',
    description: 'Limited Power of Attorney for granting specific authority to act on your behalf in defined situations',
    category: 'legal',
    downloadCount: 5670,
    fileSize: '36.9 KB',
    rating: 4.6,
    tags: ['power-of-attorney', 'legal', 'authorization', 'limited', 'agent'],
    createdAt: '2024-01-15T09:30:00Z',
    updatedAt: '2024-02-05T10:00:00Z',
    content: `
LIMITED POWER OF ATTORNEY

NOTICE: This is an important legal document. It authorizes another person to act on your behalf. Before signing, you should understand the powers you are granting and the obligations you are undertaking. Consider consulting an attorney.

PRINCIPAL INFORMATION:

I, [Full Legal Name] ("Principal")
Address: [Street Address]
[City, State, ZIP Code]
Date of Birth: [DOB]
Social Security Number (last 4 digits): XXX-XX-[####]

AGENT INFORMATION:

Hereby appoint [Full Legal Name] ("Agent" or "Attorney-in-Fact")
Address: [Street Address]
[City, State, ZIP Code]
Phone: [Phone Number]
Email: [Email Address]

SUCCESSOR AGENT (OPTIONAL):

☐ N/A - No successor agent designated

☐ If the Agent is unable or unwilling to serve, I appoint [Full Legal Name] as Successor Agent
Address: [Street Address]
[City, State, ZIP Code]
Phone: [Phone Number]
Email: [Email Address]

GRANT OF AUTHORITY:

This Limited Power of Attorney grants Agent authority to act on Principal's behalf ONLY for the specific purposes and matters described below:

1. SCOPE OF AUTHORITY

Agent is authorized to perform ONLY the following acts on behalf of Principal:

☐ Real Estate Transactions
• Buy, sell, lease, mortgage, or manage real property located at: [Property address or description]
• Execute deeds, contracts, and closing documents
• [Other specific real estate powers]

☐ Financial and Banking Matters
• Access bank accounts at [Bank name]: Account #[####]
• Make deposits and withdrawals
• Transfer funds
• Pay bills and expenses related to: [Specify]
• [Other specific financial powers]

☐ Business Operations
• Manage business operations of: [Business name]
• Sign contracts and agreements related to: [Specify]
• Hire and terminate employees
• [Other specific business powers]

☐ Tax Matters
• Prepare and file tax returns for tax year(s): [Year(s)]
• Represent Principal before [IRS/State tax authority]
• Sign tax forms and documents
• [Other specific tax powers]

☐ Legal Matters
• Represent Principal in the matter of: [Case name or description]
• Sign legal documents related to: [Specify]
• [Other specific legal powers]

☐ Insurance
• Manage insurance policies: [Policy numbers or types]
• File claims
• Change beneficiaries (if applicable)
• [Other specific insurance powers]

☐ Government Benefits
• Apply for and manage: [Social Security, Medicare, Veterans benefits, etc.]
• Sign applications and forms
• [Other specific benefits powers]

☐ Vehicle Transactions
• Buy, sell, or register vehicles: [Specify year/make/model or VIN]
• Sign title and registration documents
• [Other specific vehicle powers]

☐ Other Specific Powers
[Detailed description of any other specific, limited powers granted]

2. LIMITATIONS ON AUTHORITY

Agent does NOT have authority to:

☐ Make gifts of Principal's property (unless specifically authorized below)
☐ Create, amend, or revoke a trust
☐ Make or change beneficiary designations
☐ Create or change rights of survivorship
☐ Delegate authority to another person
☐ Sign a will or codicil on Principal's behalf
☐ Make health care decisions (separate healthcare power of attorney required)
☐ [Other limitations]

Agent's authority is limited to the specific powers granted in Section 1 above. Agent has NO general power of attorney.

3. GIFTING POWERS (IF APPLICABLE)

☐ N/A - Agent has NO authority to make gifts

☐ Agent is authorized to make gifts ONLY as follows:
• To the following individuals/organizations: [Names]
• Maximum gift per person/organization: $[Amount] per year
• Purpose: [e.g., annual exclusion gifts, charitable donations]
• Other restrictions: [Specify]

4. TERM AND EFFECTIVENESS

4.1 Effective Date
This Power of Attorney is effective:
☐ Immediately upon execution
☐ Upon the occurrence of: [Specify event or condition]
☐ On [Specific date]

4.2 Expiration
This Power of Attorney shall:
☐ Expire on [Specific date]
☐ Expire upon completion of: [Specific purpose or transaction]
☐ Continue until revoked by Principal
☐ Expire [#] months/years from the effective date

4.3 Durable or Non-Durable
☐ This is a DURABLE Power of Attorney. It remains effective if Principal becomes incapacitated.

☐ This is a NON-DURABLE Power of Attorney. It terminates if Principal becomes incapacitated.

Note: Unless specifically stated as durable, this power of attorney terminates upon Principal's incapacity under applicable state law.

5. OBLIGATIONS AND DUTIES OF AGENT

5.1 Fiduciary Duty
Agent is a fiduciary and must:
• Act in good faith and in Principal's best interests
• Act with care, competence, and diligence
• Keep Principal's property separate from Agent's property
• Avoid conflicts of interest
• Keep adequate records of all transactions

5.2 Use of Authority
Agent shall:
• Use authority only for Principal's benefit
• Make decisions consistent with Principal's known wishes
• Not use authority for Agent's own benefit unless specifically authorized
• Act within the scope of authority granted

5.3 Record Keeping
Agent shall:
• Keep accurate records of all transactions
• Keep receipts and documentation
• Provide accountings to Principal upon request
• Provide accountings to [Other person, if specified]

5.4 Reporting
Agent shall provide written reports to Principal:
☐ Upon request
☐ [Quarterly/Annually]
☐ N/A - No periodic reporting required

5.5 Compensation
Agent shall be compensated as follows:
☐ No compensation
☐ Reasonable compensation of $[Amount] or [Hourly rate]
☐ Reimbursement of reasonable expenses only
☐ As agreed between Principal and Agent

6. THIRD-PARTY RELIANCE

6.1 Acceptance
Third parties may rely on this Power of Attorney and are not responsible for determining:
• Whether the power has been revoked
• Whether the Principal is alive or competent
• Whether the Agent is acting properly within the scope of authority

Except: Third parties have actual knowledge of revocation or termination.

6.2 Protection
Third parties who accept and act in reliance on this Power of Attorney without actual knowledge that it has been revoked or terminated shall be fully protected.

6.3 Certification
Agent may execute an affidavit or certification stating:
• Principal is alive and has not revoked this Power of Attorney
• Agent's authority is in full force and effect
• Agent is acting within the scope of authority

7. NOMINATION OF GUARDIAN (OPTIONAL)

☐ N/A

☐ If a court proceeding is initiated to appoint a guardian of the person or estate for Principal, Principal nominates [Name] to serve as guardian.

This nomination does not give Agent any authority over healthcare decisions unless separately authorized.

8. REVOCATION

8.1 Right to Revoke
Principal may revoke this Power of Attorney at any time while competent by:
• Executing a written revocation
• Executing a new Power of Attorney
• Notifying Agent in writing
• Destroying all original copies

8.2 Notice of Revocation
Upon revocation, Principal should provide notice to:
• Agent
• All third parties who have relied on this Power of Attorney
• Any institution or person holding a copy

8.3 Automatic Termination
This Power of Attorney automatically terminates upon:
• Principal's death
• Expiration date specified in Section 4.2
• Completion of the specific purpose
• Principal's incapacity (if non-durable)
• Court appointment of a guardian or conservator (in some jurisdictions)

9. GOVERNING LAW

This Power of Attorney shall be governed by and construed in accordance with the laws of the State of [State].

10. ACKNOWLEDGMENT AND ACCEPTANCE BY AGENT

Agent acknowledges:
• Receipt of this Power of Attorney
• Understanding of the fiduciary duties and obligations
• Agreement to act in Principal's best interests
• Agreement to keep adequate records
• Authority is limited to the specific powers granted herein

11. SEVERABILITY

If any provision of this Power of Attorney is held invalid, the remaining provisions shall continue in effect to the extent possible.

12. PHOTOCOPIES

A photocopy or electronic copy of this Power of Attorney shall have the same effect as the original.


PRINCIPAL'S SIGNATURE:

I have read this Limited Power of Attorney and understand its contents. I am of sound mind and am executing this document voluntarily, without undue influence.

Principal Signature: _________________________
Printed Name: [Full Legal Name]
Date: _____________________________


AGENT'S ACKNOWLEDGMENT AND ACCEPTANCE:

I acknowledge that I have read this Limited Power of Attorney, understand my duties and obligations, and agree to act as Agent for the Principal under the terms set forth herein.

Agent Signature: _________________________
Printed Name: [Full Legal Name]
Date: _____________________________


WITNESSES (IF REQUIRED BY STATE LAW):

Witness 1 Signature: _________________________
Printed Name: [Full Legal Name]
Address: [Address]
Date: _____________________________

Witness 2 Signature: _________________________
Printed Name: [Full Legal Name]
Address: [Address]
Date: _____________________________


NOTARY ACKNOWLEDGMENT:

STATE OF [STATE]
COUNTY OF [COUNTY]

On this _____ day of ______________, 20___, before me, the undersigned notary public, personally appeared [Principal's Name], proved to me through satisfactory evidence of identification, which was [type of ID], to be the person whose name is subscribed to the within instrument and acknowledged to me that [he/she/they] executed the same in [his/her/their] authorized capacity, and that by [his/her/their] signature on the instrument, the person, or the entity upon behalf of which the person acted, executed the instrument.

Witness my hand and official seal.


Notary Public Signature: _________________________
Printed Name: [Notary Name]
My Commission Expires: _____________________

[NOTARY SEAL]


IMPORTANT NOTICES:

1. This form may not be suitable for all situations. Consult an attorney if you have questions.

2. Requirements for Powers of Attorney vary by state. Ensure this document complies with your state's laws.

3. Some institutions may require their own forms or additional documentation.

4. Keep the original in a safe place. Provide certified copies to Agent and relevant third parties.

5. Review and update this Power of Attorney periodically to ensure it reflects your current wishes.

6. This Power of Attorney does NOT authorize healthcare decisions. A separate Healthcare Power of Attorney or Advance Directive is needed for medical decisions.
    `
  }
];

// Add legal & admin templates to the main templates array
templates.push(...legalAdminTemplates);

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