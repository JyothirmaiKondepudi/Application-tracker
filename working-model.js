const HEADERS = ['Company Name', 'Job Title', 'Application URL', 'Application Status', 'Date Applied', 'Last Updated'];
const SHEET_NAME = 'Job Applications'; // Name of your Google Sheet


// Job application patterns to identify new applications
const JOB_APPLICATION_PATTERNS = [
  "thank you for applying",
  "thank you for your application", 
  "application received",
  "we have received your application",
  "we've received your application",
  "application confirmation",
  "your application has been submitted",
  "application submitted successfully",
  "thanks for applying",
  "application complete"
];

// Status update patterns for different stages
const STATUS_PATTERNS = {
  'Assessment': [
    'online assessment',
    'coding challenge',
    'technical assessment',
    'complete the assessment',
    'skills assessment',
    'programming test'
  ],
  'Interview': [
    'interview invitation',
    'schedule an interview',
    'interview scheduled',
    'phone interview',
    'video interview',
    'technical interview',
    'behavioral interview'
  ],
  'HR Round': [
    'hr interview',
    'hr round',
    'human resources',
    'final interview',
    'culture fit interview'
  ],
  'Offer': [
    'job offer',
    'offer letter',
    'congratulations',
    'pleased to offer',
    'extend an offer',
    'offer of employment'
  ],
  'Rejected': [
    'unfortunately',
    'not selected',
    'decided to move forward with other candidates',
    'will not be moving forward',
    'thank you for your interest, however',
    'regret to inform',
    'not a match at this time'
  ]
};

// STEP 1: Run this function first to set up everything
function setupJobTracker() {
  try {
    setupSpreadsheet();
    createTimerTrigger();
    Logger.log('✅ Job tracker setup complete!');
    Logger.log('🔄 Email checking will start in 15 minutes');
    
    // Run a test to check recent emails immediately
    Logger.log('🧪 Running initial email check...');
    checkRecentEmails();
    
  } catch (error) {
    Logger.log('❌ Setup failed: ' + error.toString());
  }
}

// Set up the spreadsheet with headers
function setupSpreadsheet() {
  let ss;
  
  try {
    // Try to get active spreadsheet first
    ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('📊 Using active spreadsheet: ' + ss.getName());
  } catch (error) {
    // If no active spreadsheet, create a new one
    ss = SpreadsheetApp.create('Job Application Tracker');
    Logger.log('📊 Created new spreadsheet: ' + ss.getUrl());
  }
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    Logger.log('📝 Created new sheet: ' + SHEET_NAME);
  }
  
  // Add headers if they don't exist or if first row is empty
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = firstRow.some(cell => cell !== '');
  
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    Logger.log('✅ Headers added to spreadsheet');
  }
  
  Logger.log('✅ Spreadsheet setup complete');
}

// Create a time-based trigger to check emails periodically
function createTimerTrigger() {
  // Delete existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkRecentEmails') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  
  if (deletedCount > 0) {
    Logger.log(`🗑️ Deleted ${deletedCount} existing triggers`);
  }
  
  // Create time-based trigger (runs every 15 minutes)
  ScriptApp.newTrigger('checkRecentEmails')
    .timeBased()
    .everyMinutes(15)
    .create();
  
  Logger.log('⏰ Time-based trigger created - will check emails every 15 minutes');
}

// STEP 2: This function runs automatically every 15 minutes
function checkRecentEmails() {
  try {
    Logger.log('🔍 Checking recent emails...');
    
    // Get emails from last 20 minutes (to ensure we don't miss any)
    const twentyMinutesAgo = new Date(Date.now() - 20 * 60 * 1000);
    const searchQuery = `after:${Math.floor(twentyMinutesAgo.getTime()/1000)}`;
    const threads = GmailApp.search(searchQuery, 0, 50);
    
    Logger.log(`📧 Found ${threads.length} email threads to check`);
    
    let newApplications = 0;
    let statusUpdates = 0;
    
    for (let thread of threads) {
      const messages = thread.getMessages();
      
      for (let message of messages) {
        if (message.getDate() > twentyMinutesAgo) {
          const subject = message.getSubject().toLowerCase();
          const body = message.getPlainBody().toLowerCase();
          
          if (isJobApplicationEmail(subject, body)) {
            const jobData = extractJobInfo(message);
            if (!applicationExists(jobData)) {
              addJobApplication(jobData);
              newApplications++;
              Logger.log(`✅ New job application added: ${jobData.companyName} - ${jobData.jobTitle}`);
            }
          }
          else if (isStatusUpdateEmail(subject, body)) {
            const updateInfo = extractStatusUpdate(message);
            if (updateJobStatus(updateInfo)) {
              statusUpdates++;
              Logger.log(`🔄 Status updated: ${updateInfo.companyName} -> ${updateInfo.status}`);
            }
          }
        }
      }
    }
    
    Logger.log(`📊 Summary: ${newApplications} new applications, ${statusUpdates} status updates`);
    
  } catch (error) {
    Logger.log(`❌ Error checking recent emails: ${error.toString()}`);
  }
}



// STEP 3: Run this manually to scan your existing emails for job applications
function scanExistingEmails() {
  try {
    Logger.log('🔍 Scanning existing emails for job applications...');
    
    // Search for emails with job-related keywords
    const searchQueries = [
      'subject:(application received OR thank you for applying OR application confirmation)',
      'subject:(interview OR assessment OR coding challenge)',
      'from:(noreply OR careers OR jobs OR recruiting)'
    ];
    
    let totalFound = 0;
    
    for (let query of searchQueries) {
      const threads = getAllThreads(query);
      Logger.log(`📧 Checking ${threads.length} threads for query: ${query}`);
      
      for (let thread of threads) {
        const messages = thread.getMessages();
        
        for (let message of messages) {
          const subject = message.getSubject().toLowerCase();
          const body = message.getPlainBody().toLowerCase();
          
          if (isJobApplicationEmail(subject, body)) {
            const jobData = extractJobInfo(message);
            if (!applicationExists(jobData)) {
              addJobApplication(jobData);
              totalFound++;
              Logger.log(`✅ Added: ${jobData.companyName} - ${jobData.jobTitle}`);
            }
          }
        }
      }
    }
    
    Logger.log(`🎉 Scan complete! Found and added ${totalFound} job applications`);
    
  } catch (error) {
    Logger.log(`❌ Error scanning existing emails: ${error.toString()}`);
  }
}
/**
 * Runs scanExistingEmails() in date‐range “buckets,” e.g. one bucket per quarter or per year.
 * You’ll need to decide how large each date chunk should be so that each search(...) call
 * returns fewer than ~500 threads. In this example, we break it into 6-month intervals.
 */

function scanOlderEmailsInBuckets() {
  const DATE_BUCKETS = [
    // [ “YYYY-MM-DD” of bucketStart, “YYYY-MM-DD” of bucketEnd ]
    [ "2025-01-01", "2025-06-30" ]
    
  ];
  
  let totalFound = 0;
  
  // The same search-patterns you use in scanExistingEmails(),
  // but we’ll append “after:YYYY/MM/DD before:YYYY/MM/DD” to each.
  const baseQueries = [
    'subject:(application received OR thank you for applying OR application confirmation)',
    'subject:(interview OR assessment OR coding challenge)',
    'from:(noreply OR careers OR jobs OR recruiting)'
  ];
  
  for (let [afterDate, beforeDate] of DATE_BUCKETS) {
    for (let base of baseQueries) {
      // Build a date-restricted query
      // – “in:anywhere” ensures we search beyond “in:inbox” if needed
      // – note the format must be YYYY/MM/DD for Gmail search
      const dateQuery = `${base} in:anywhere after:${afterDate.replace(/-/g, "/")} before:${beforeDate.replace(/-/g, "/")}`;
      
      // Now page through *all* threads in that date range
      const threads = getAllThreads(dateQuery);
      Logger.log(`📧 Buckets [${afterDate} → ${beforeDate}] has ${threads.length} threads for query: ${base}`);
      
      for (let thread of threads) {
        const messages = thread.getMessages();
        for (let message of messages) {
          const subject = message.getSubject().toLowerCase();
          const body = message.getPlainBody().toLowerCase();
          
          if (isJobApplicationEmail(subject, body)) {
            const jobData = extractJobInfo(message);
            if (!applicationExists(jobData)) {
              addJobApplication(jobData);
              totalFound++;
            }
          }
        }
      }
    }
  }
  
  Logger.log(`🎉 DONE: scanOlderEmailsInBuckets() added ${totalFound} older job applications.`);
}


/**
 * getAllThreads(query)
 *
 * Exactly the same helper from before—pages through GmailApp.search(...),
 * 100 at a time, until no more threads remain.
 */
function getAllThreads(query) {
  const BATCH_SIZE = 100;
  let allThreads = [];
  let start = 0;
  while (true) {
    const batch = GmailApp.search(query, start, BATCH_SIZE);
    if (!batch || batch.length === 0) break;
    allThreads = allThreads.concat(batch);
    if (batch.length < BATCH_SIZE) break;
    start += BATCH_SIZE;
  }
  return allThreads;
}

// --------------------------------------------------------------------------------
// The following helper functions must already exist in your script from last time:
//   • isJobApplicationEmail(subject, body)
//   • extractJobInfo(message)
//   • applicationExists(jobData)
//   • addJobApplication(jobData)
//   • (…and all of their dependencies, e.g. extractCompanyName, extractJobTitle, etc.)
// --------------------------------------------------------------------------------


// Helper function to check if application already exists
function applicationExists(jobData) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return false; // Only headers or empty
    
    for (let i = 1; i < data.length; i++) {
      const existingCompany = (data[i][0] || '').toString().toLowerCase();
      const existingJob = (data[i][1] || '').toString().toLowerCase();
      const newCompany = jobData.companyName.toLowerCase();
      const newJob = jobData.jobTitle.toLowerCase();
      
      if (existingCompany === newCompany && existingJob === newJob) {
        return true;
      }
    }
    return false;
    
  } catch (error) {
    Logger.log(`❌ Error checking if application exists: ${error.toString()}`);
    return false;
  }
}

// Get or create the sheet
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(`📊 Working with spreadsheet: ${ss.getName()}`);
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log(`📝 Sheet '${SHEET_NAME}' not found, creating new one`);
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    Logger.log(`✅ Created sheet '${SHEET_NAME}' with headers`);
  } else {
    Logger.log(`📝 Found existing sheet: ${SHEET_NAME}`);
  }
  
  return sheet;
}

// Check if email is a job application confirmation
function isJobApplicationEmail(subject, body) {
  return JOB_APPLICATION_PATTERNS.some(pattern => 
    subject.includes(pattern) || body.includes(pattern)
  );
}

// Check if email is a status update
function isStatusUpdateEmail(subject, body) {
  const allStatusPatterns = Object.values(STATUS_PATTERNS).flat();
  return allStatusPatterns.some(pattern => 
    subject.includes(pattern) || body.includes(pattern)
  );
}

// Extract job information from application email
function extractJobInfo(message) {
  const subject = message.getSubject();
  const body = message.getPlainBody();
  const sender = message.getFrom();
  const date = message.getDate();
  
  return {
    companyName: extractCompanyName(subject, body, sender),
    jobTitle: extractJobTitle(subject, body),
    applicationURL: extractJobURL(body),
    applicationStatus: 'Applied',
    dateApplied: date,
    lastUpdated: date
  };
}

// Extract company name from email
function extractCompanyName(subject, body, sender) {
 // — first catch “…– Position at Company” in the subject
  const dashCompany = subject.match(/-\s*.+?\s+at\s+([^\s-–—]+)/i);
  if (dashCompany && dashCompany[1]) {
    return capitalizeWords(dashCompany[1].trim());
}

  // then fall back to domain or existing logic

  const emailMatch = sender.match(/@([^.]+)\.(?:com|org|net|io|co)/i);
  if (emailMatch && !isGenericDomain(emailMatch[1])) {
    return capitalizeWords(emailMatch[1]);
  }
  
  // Common patterns in subject lines
  const companyPatterns = [
    /application.*?(?:at|to|with)\s+([A-Za-z0-9\s&]+?)(?:\s*[-,]|$)/i,
    /([A-Za-z0-9\s&]+?)\s+(?:team|careers|jobs|application)/i,
    /position.*?(?:at|with)\s+([A-Za-z0-9\s&]+?)(?:\s*[-,]|$)/i,
    /thank you for applying.*?(?:at|to|with)\s+([A-Za-z0-9\s&]+?)(?:\s*[-,]|$)/i
  ];
  
  for (let pattern of companyPatterns) {
    const match = subject.match(pattern);
    if (match && match[1].trim().length > 1) {
      return capitalizeWords(match[1].trim());
    }
  }
  
  // Look in email body
  const bodyPatterns = [
    /company:\s*([^\n\r]+)/i,
    /employer:\s*([^\n\r]+)/i,
    /from\s+([A-Za-z0-9\s&]+?)(?:\s+careers|\s+team)/i
  ];
  
  for (let pattern of bodyPatterns) {
    const match = body.match(pattern);
    if (match && match[1].trim().length > 1) {
      return capitalizeWords(match[1].trim());
    }
  }
  
  // Fallback: use sender domain
  const domainFallback = sender.match(/@([^.]+)/);
  if (domainFallback) {
    return capitalizeWords(domainFallback[1]);
  }
  
  return 'Unknown Company';
}

// Extract job title from email
function extractJobTitle(subject, body) {
  const dashTitle = subject.match(/-\s*(.+?)\s+at\s+[^\s-–—]+/i);
  if (dashTitle && dashTitle[1]) {
    return capitalizeWords(dashTitle[1].trim());
  }
  // Common patterns in subject lines
  const titlePatterns = [
    /application.*?for\s+([^-,\n\r]+?)(?:\s+(?:position|role|at|application)|$)/i,
    /([^-,\n\r]+?)\s+(?:application|position)\s+(?:received|confirmed)/i,
    /position:\s*([^-,\n\r]+)/i,
    /role:\s*([^-,\n\r]+)/i,
    /applied.*?(?:for|to)\s+([^-,\n\r]+?)(?:\s+(?:position|role|at)|$)/i
  ];
  
  for (let pattern of titlePatterns) {
    const match = subject.match(pattern) || body.match(pattern);
    if (match && match[1].trim().length > 1) {
      return capitalizeWords(match[1].trim());
    }
  }
  
  return 'Unknown Position';
}

// Extract job URL from email body
function extractJobURL(body) {
  const urlPatterns = [
    /(https?:\/\/[^\s]+(?:job|career|apply|position|linkedin|indeed)[^\s]*)/i,
    /(https?:\/\/(?:www\.)?linkedin\.com\/jobs\/[^\s]+)/i,
    /(https?:\/\/[^\s]*indeed\.com[^\s]*)/i,
    /(https?:\/\/[^\s]*glassdoor\.com[^\s]*)/i
  ];
  
  for (let pattern of urlPatterns) {
    const match = body.match(pattern);
    if (match) {
      return match[1].replace(/[.,;)]$/, ''); // Remove trailing punctuation
    }
  }
  
  return '';
}

// Extract status update information
function extractStatusUpdate(message) {
  const subject = message.getSubject().toLowerCase();
  const body = message.getPlainBody().toLowerCase();
  const sender = message.getFrom();
  
  // Determine new status
  let newStatus = '';
  for (let [status, patterns] of Object.entries(STATUS_PATTERNS)) {
    if (patterns.some(pattern => subject.includes(pattern) || body.includes(pattern))) {
      newStatus = status;
      break;
    }
  }
  
  return {
    sender: sender,
    status: newStatus,
    companyName: extractCompanyName(message.getSubject(), message.getPlainBody(), sender),
    date: message.getDate()
  };
}

// Add new job application to spreadsheet
function addJobApplication(jobData) {
  try {
    const sheet = getOrCreateSheet();
    Logger.log(`📝 Adding to sheet: ${sheet.getName()}`);
    
    const row = [
      jobData.companyName,
      jobData.jobTitle,
      jobData.applicationURL,
      jobData.applicationStatus,
      jobData.dateApplied,
      jobData.lastUpdated
    ];
    
    Logger.log(`📊 Row data: ${JSON.stringify(row)}`);
    
    // Get the next row number
    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;
    Logger.log(`📍 Adding to row ${nextRow}`);
    
    // Add the row
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    
    // Verify it was added
    const addedData = sheet.getRange(nextRow, 1, 1, row.length).getValues()[0];
    Logger.log(`✅ Verified data added: ${JSON.stringify(addedData)}`);
    Logger.log(`✅ Added application: ${jobData.companyName} - ${jobData.jobTitle}`);
    
  } catch (error) {
    Logger.log(`❌ Error adding job application: ${error.toString()}`);
    Logger.log(`❌ Stack trace: ${error.stack}`);
  }
}

// Update existing job application status
function updateJobStatus(updateInfo) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return false; // Only headers or empty
    
    // Find matching row by company name
    for (let i = 1; i < data.length; i++) {
      const companyName = (data[i][0] || '').toString().toLowerCase();
      const searchCompany = updateInfo.companyName.toLowerCase();
      
      // Simple fuzzy matching
      if (companyName.includes(searchCompany) || searchCompany.includes(companyName)) {
        // Update status and last updated date
        sheet.getRange(i + 1, 4).setValue(updateInfo.status); // Status column
        sheet.getRange(i + 1, 6).setValue(updateInfo.date); // Last updated column
        
        Logger.log(`🔄 Updated ${companyName} status to ${updateInfo.status}`);
        return true;
      }
    }
    
    Logger.log(`❌ Could not find matching application for: ${updateInfo.companyName}`);
    return false;
    
  } catch (error) {
    Logger.log(`❌ Error updating job status: ${error.toString()}`);
    return false;
  }
}

// Utility function to check if domain is generic
function isGenericDomain(domain) {
  const genericDomains = [
    'workday', 'greenhouse', 'lever', 'bamboohr', 'jobvite', 
    'smartrecruiters', 'gmail', 'outlook', 'yahoo', 'noreply',
    'no-reply', 'donotreply', 'automated'
  ];
  return genericDomains.some(generic => domain.toLowerCase().includes(generic));
}

// Utility function to capitalize words
function capitalizeWords(str) {
  return str.replace(/\b\w+/g, word => 
    word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
  );
}

// TESTING AND DEBUGGING FUNCTIONS

// Test function with sample data
function testWithSampleData() {
  Logger.log('🧪 Testing with sample data...');
  
  const sampleJobData = {
    companyName: 'Test Company',
    jobTitle: 'Software Engineer',
    applicationURL: 'https://example.com/job',
    applicationStatus: 'Applied',
    dateApplied: new Date(),
    lastUpdated: new Date()
  };
  
  // Test if application exists (should return false)
  const exists = applicationExists(sampleJobData);
  Logger.log(`Application exists: ${exists}`);
  
  // Add the sample application
  addJobApplication(sampleJobData);
  Logger.log('✅ Sample application added');
}

// Test extraction functions
function testExtraction() {
  Logger.log('🧪 Testing extraction functions...');
  
  const testSubject = "Thank you for your application - Software Engineer at Google";
  const testBody = "Thank you for applying to the Software Engineer position at Google. You can view your application at https://careers.google.com/jobs/123";
  const testSender = "noreply@google.com";
  
  Logger.log("Company:", extractCompanyName(testSubject, testBody, testSender));
  Logger.log("Job Title:", extractJobTitle(testSubject, testBody));
  Logger.log("URL:", extractJobURL(testBody));
}

// Debug function to check current triggers
function debugTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`📋 Found ${triggers.length} triggers:`);
  
  triggers.forEach((trigger, index) => {
    Logger.log(`${index + 1}. Function: ${trigger.getHandlerFunction()}`);
    Logger.log(`   Type: ${trigger.getTriggerSource()}`);
    Logger.log(`   ID: ${trigger.getUniqueId()}`);
  });
}

// DEBUGGING: Check what sheets exist and their content
function debugSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`📊 Active Spreadsheet: ${ss.getName()}`);
    Logger.log(`📊 Spreadsheet URL: ${ss.getUrl()}`);
    
    const sheets = ss.getSheets();
    Logger.log(`📝 Found ${sheets.length} sheets:`);
    
    sheets.forEach((sheet, index) => {
      Logger.log(`${index + 1}. Sheet name: "${sheet.getName()}"`);
      Logger.log(`   Rows: ${sheet.getLastRow()}, Columns: ${sheet.getLastColumn()}`);
      
      if (sheet.getLastRow() > 0) {
        const data = sheet.getRange(1, 1, Math.min(3, sheet.getLastRow()), sheet.getLastColumn()).getValues();
        Logger.log(`   First few rows: ${JSON.stringify(data)}`);
      }
    });
    
    // Specifically check for our target sheet
    const targetSheet = ss.getSheetByName(SHEET_NAME);
    if (targetSheet) {
      Logger.log(`✅ Target sheet '${SHEET_NAME}' exists`);
      Logger.log(`📊 Sheet dimensions: ${targetSheet.getLastRow()} rows, ${targetSheet.getLastColumn()} columns`);
    } else {
      Logger.log(`❌ Target sheet '${SHEET_NAME}' NOT found`);
    }
    
  } catch (error) {
    Logger.log(`❌ Error debugging sheets: ${error.toString()}`);
  }
}

// DEBUGGING: Add a simple test row directly
function debugAddTestRow() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`📊 Active Spreadsheet: ${ss.getName()}`);
    
    // Get the first sheet (regardless of name)
    const sheet = ss.getSheets()[0];
    Logger.log(`📝 Using sheet: ${sheet.getName()}`);
    
    // Add a test row
    const testRow = ['TEST COMPANY', 'TEST POSITION', 'http://test.com', 'Applied', new Date(), new Date()];
    const nextRow = sheet.getLastRow() + 1;
    
    Logger.log(`📍 Adding test row to row ${nextRow}`);
    sheet.getRange(nextRow, 1, 1, testRow.length).setValues([testRow]);
    
    Logger.log(`✅ Test row added successfully`);
    
    // Verify
    const addedData = sheet.getRange(nextRow, 1, 1, testRow.length).getValues()[0];
    Logger.log(`✅ Verified: ${JSON.stringify(addedData)}`);
    
  } catch (error) {
    Logger.log(`❌ Error adding test row: ${error.toString()}`);
    Logger.log(`❌ Stack: ${error.stack}`);
  }
}