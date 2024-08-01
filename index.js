const { Sequelize, DataTypes } = require('sequelize');
const { faker } = require('@faker-js/faker');
const xlsx = require('xlsx'); // Import the xlsx package

// Create a new Sequelize instance
const sequelize = new Sequelize({
  dialect: 'sqlite',
  storage: 'database.sqlite', // This will create a new SQLite database file
});

// Define the Students table
const Student = sequelize.define('Student', {
  OSIS_Number: {
    type: DataTypes.STRING(9),
    primaryKey: true,
  },
  First_Name: { type: DataTypes.STRING, allowNull: false },
  Middle_Name: { type: DataTypes.STRING },
  Last_Name: { type: DataTypes.STRING, allowNull: false },
  Date_of_Birth: { type: DataTypes.DATE, allowNull: false },
  Gender: { type: DataTypes.STRING, allowNull: false },
  Street_Address: { type: DataTypes.STRING, allowNull: false },
  Apt: { type: DataTypes.STRING },
  City: { type: DataTypes.STRING, allowNull: false },
  State: { type: DataTypes.STRING, allowNull: false },
  Zip_Code: { type: DataTypes.STRING, allowNull: false },
  Current_Grade: { type: DataTypes.INTEGER, allowNull: false },
  Current_School: { type: DataTypes.STRING, allowNull: false },
  Current_School_Address: { type: DataTypes.STRING, allowNull: false },
  Home_Language: { type: DataTypes.STRING, allowNull: false },
  Special_Education_Services: { type: DataTypes.BOOLEAN, allowNull: false },
  Family_ID: { type: DataTypes.INTEGER, allowNull: false },
  Takes_Bus: { type: DataTypes.BOOLEAN, allowNull: false },
  Bus_ID: { type: DataTypes.INTEGER },
  Dismissal_Method: { type: DataTypes.STRING, allowNull: false },
  Economically_Disadvantaged: { type: DataTypes.BOOLEAN, allowNull: false },
  Ethnicity: { type: DataTypes.STRING, allowNull: false }, // Add Ethnicity field
});

// Define the Families table
const Family = sequelize.define('Family', {
  Family_ID: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
  Primary_Contact_Name: { type: DataTypes.STRING, allowNull: false },
  Primary_Contact_Phone: { type: DataTypes.STRING, allowNull: false },
  Primary_Contact_Email: { type: DataTypes.STRING, allowNull: false },
  Home_Language: { type: DataTypes.STRING, allowNull: false },
  Address: { type: DataTypes.STRING, allowNull: false },
  Secondary_Contact_Name: { type: DataTypes.STRING, allowNull: false },
  Secondary_Contact_Phone: { type: DataTypes.STRING, allowNull: false },
  Secondary_Contact_Email: { type: DataTypes.STRING, allowNull: false },
  Relationship_to_Student: { type: DataTypes.STRING, allowNull: false },
});

// Define the Survey_Submissions table
const SurveySubmission = sequelize.define('SurveySubmission', {
  Submission_ID: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true,
  },
  OSIS_Number: { type: DataTypes.STRING(9), allowNull: false }, // Add OSIS_Number field
  Submission_Status: { type: DataTypes.BOOLEAN, allowNull: false },
  Submission_Date: { type: DataTypes.DATE },
  Follow_Up_Attempts: { type: DataTypes.INTEGER, allowNull: false },
  Last_Contacted: { type: DataTypes.DATE },
  Follow_Up_Notes: { type: DataTypes.TEXT },
});

// Define the Contacts table
const Contact = sequelize.define('Contact', {
  Contact_ID: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true,
  },
  Family_ID: { type: DataTypes.INTEGER, allowNull: false },
  Contact_Date: { type: DataTypes.DATE, allowNull: false },
  Contact_Method: { type: DataTypes.STRING, allowNull: false },
  Contact_Notes: { type: DataTypes.TEXT, allowNull: false },
});

// Define the Buses table
const Bus = sequelize.define('Bus', {
  Bus_ID: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
  Bus_Number: { type: DataTypes.STRING, allowNull: false },
  Route: { type: DataTypes.STRING, allowNull: false },
});

// Define the Emergency_Contacts table
const EmergencyContact = sequelize.define('EmergencyContact', {
  Contact_ID: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true,
  },
  OSIS_Number: { type: DataTypes.STRING(9), allowNull: false }, // Add OSIS_Number field
  Contact_Name: { type: DataTypes.STRING, allowNull: false },
  Contact_Relation: { type: DataTypes.STRING, allowNull: false },
  Contact_Phone: { type: DataTypes.STRING, allowNull: false },
});

// Set up relationships
Student.belongsTo(Family, { foreignKey: 'Family_ID' });
Student.belongsTo(Bus, { foreignKey: 'Bus_ID' });
SurveySubmission.belongsTo(Student, {
  foreignKey: 'OSIS_Number',
  targetKey: 'OSIS_Number',
});
Contact.belongsTo(Family, { foreignKey: 'Family_ID' });
EmergencyContact.belongsTo(Student, {
  foreignKey: 'OSIS_Number',
  targetKey: 'OSIS_Number',
});

// Function to get a random integer between min and max (inclusive)
function getRandomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// Function to create random students and families
async function createSampleData() {
  const ethnicDistribution = [
    { ethnicity: 'Hispanic or Latino', percentage: 94 },
    { ethnicity: 'Black or African American', percentage: 3 },
    { ethnicity: 'White', percentage: 1 },
    {
      ethnicity: 'Asian or Native Hawaiian/Other Pacific Islander',
      percentage: 1,
    },
    { ethnicity: 'Multiracial', percentage: 1 },
  ];

  const totalStudents = 100;
  const families = [];
  const students = [];
  const totalBuses = 10; // Generate a realistic number of buses
  const buses = [];

  // Generate random buses
  for (let i = 0; i < totalBuses; i++) {
    const bus = await Bus.create({
      Bus_Number: `Bus-${i + 1}`,
      Route: `Route-${i + 1}`,
    });
    buses.push(bus);
  }

  // Generate random families
  for (let i = 0; i < totalStudents / 2; i++) {
    const family = await Family.create({
      Primary_Contact_Name: faker.person.fullName(),
      Primary_Contact_Phone: faker.phone.number(),
      Primary_Contact_Email: faker.internet.email(),
      Home_Language: 'English', // Default to English; will adjust for ELLs later
      Address: `${faker.location.streetAddress()}, New York, NY ${faker.location.zipCode()}`,
      Secondary_Contact_Name: faker.person.fullName(),
      Secondary_Contact_Phone: faker.phone.number(),
      Secondary_Contact_Email: faker.internet.email(),
      Relationship_to_Student: faker.helpers.arrayElement([
        'Mother',
        'Father',
        'Guardian',
      ]),
    });

    families.push(family);
  }

  // Define number of ELLs and their languages
  const totalELLs = Math.round(totalStudents * 0.25); // 25% of students are ELLs
  let spanishELLs = Math.round(totalELLs * 0.98); // 98% of ELLs speak Spanish
  let arabicELLs = Math.round(totalELLs * 0.01); // 1% of ELLs speak Arabic
  let russianELLs = Math.round(totalELLs * 0.01); // 1% of ELLs speak Russian

  // Define number of students with disabilities
  let totalStudentsWithDisabilities = Math.round(totalStudents * 0.213); // 21.3% of students with disabilities

  // Define number of economically disadvantaged students
  let totalEconomicallyDisadvantaged = Math.round(totalStudents * 0.901); // 90.1% economically disadvantaged

  // Define number of homeless students
  let totalHomeless = Math.round(totalStudents * 0.02); // 2% homeless students

  // Generate random students with ethnic distribution
  for (let i = 0; i < totalStudents; i++) {
    const ethnicity = faker.helpers.arrayElement(
      ethnicDistribution.flatMap((ethnicity) =>
        Array(ethnicity.percentage).fill(ethnicity.ethnicity)
      )
    );

    const family = families[getRandomInt(0, families.length - 1)];

    // Determine home language
    let homeLanguage = 'English';
    if (spanishELLs > 0) {
      homeLanguage = 'Spanish';
      spanishELLs--;
    } else if (arabicELLs > 0) {
      homeLanguage = 'Arabic';
      arabicELLs--;
    } else if (russianELLs > 0) {
      homeLanguage = 'Russian';
      russianELLs--;
    }

    // Update family language if ELL
    if (homeLanguage !== 'English') {
      family.Home_Language = homeLanguage;
      await family.save();
    }

    // Determine special education services
    let specialEducationServices = false;
    if (totalStudentsWithDisabilities > 0) {
      specialEducationServices = true;
      totalStudentsWithDisabilities--;
    }

    // Determine economic disadvantage
    let economicallyDisadvantaged = false;
    if (totalEconomicallyDisadvantaged > 0) {
      economicallyDisadvantaged = true;
      totalEconomicallyDisadvantaged--;
    }

    // Determine if the student is homeless
    let isHomeless = false;
    if (totalHomeless > 0) {
      isHomeless = true;
      totalHomeless--;
    }

    const student = await Student.create({
      OSIS_Number: faker.number
        .int({ min: 100000000, max: 999999999 })
        .toString(),
      First_Name: faker.person.firstName(),
      Middle_Name: faker.person.firstName(),
      Last_Name: faker.person.lastName(),
      Date_of_Birth: faker.date.between('2005-01-01', '2012-12-31'),
      Gender: faker.helpers.arrayElement(['Male', 'Female']),
      Street_Address: isHomeless ? 'Homeless' : family.Address.split(',')[0],
      Apt: isHomeless ? '' : faker.helpers.arrayElement(['1A', '2B', '3C']),
      City: isHomeless ? '' : family.Address.split(',')[1].trim(),
      State: isHomeless
        ? ''
        : family.Address.split(',')[2].trim().split(' ')[0],
      Zip_Code: isHomeless
        ? ''
        : family.Address.split(',')[2].trim().split(' ')[1],
      Current_Grade: getRandomInt(1, 8),
      Current_School: 'TEP',
      Current_School_Address: faker.location.streetAddress(),
      Home_Language: homeLanguage,
      Special_Education_Services: specialEducationServices,
      Family_ID: family.Family_ID,
      Takes_Bus: faker.datatype.boolean(),
      Bus_ID: faker.helpers.arrayElement(buses).Bus_ID,
      Dismissal_Method: faker.helpers.arrayElement(['Parent Pick Up', 'Bus']),
      Economically_Disadvantaged: economicallyDisadvantaged,
      Ethnicity: ethnicity, // Set Ethnicity
    });

    students.push(student);

    // Create a survey submission for each student
    await SurveySubmission.create({
      OSIS_Number: student.OSIS_Number,
      Submission_Status: faker.datatype.boolean(),
      Submission_Date: faker.date.recent(),
      Follow_Up_Attempts: getRandomInt(0, 3),
      Last_Contacted: faker.date.recent(),
      Follow_Up_Notes: faker.lorem.sentence(),
    });
  }

  // Generate random emergency contacts for each student
  for (const student of students) {
    await EmergencyContact.create({
      OSIS_Number: student.OSIS_Number,
      Contact_Name: faker.person.fullName(),
      Contact_Relation: faker.helpers.arrayElement([
        'Mother',
        'Father',
        'Guardian',
      ]),
      Contact_Phone: faker.phone.number(),
    });
  }

  // Generate random contact logs for each family
  for (const family of families) {
    await Contact.create({
      Family_ID: family.Family_ID,
      Contact_Date: faker.date.recent(),
      Contact_Method: faker.helpers.arrayElement([
        'Phone',
        'Email',
        'In-Person',
      ]),
      Contact_Notes: faker.lorem.sentence(),
    });
  }
}

// Function to export tables to a single Excel file with multiple sheets
async function exportToExcel() {
  const tables = [
    { model: Student, name: 'Students' },
    { model: Family, name: 'Families' },
    { model: SurveySubmission, name: 'Survey_Submissions' },
    { model: Contact, name: 'Contacts' },
    { model: Bus, name: 'Buses' },
    { model: EmergencyContact, name: 'Emergency_Contacts' },
  ];

  const workbook = xlsx.utils.book_new();

  let totalFamilies = 0;
  let completedFamilies = 0;
  let familyParticipation = {};

  for (const table of tables) {
    let data;
    if (table.name === 'Survey_Submissions') {
      // Perform join with Student to include Family_ID in Survey_Submissions
      data = await table.model.findAll({
        include: [
          {
            model: Student,
            attributes: ['Family_ID'],
          },
        ],
        raw: true,
      });

      // Track families and their completion status
      const familyCompletionStatus = {};
      data.forEach((entry) => {
        const familyID = entry['Student.Family_ID'];
        if (!familyCompletionStatus[familyID]) {
          familyCompletionStatus[familyID] = false;
        }
        if (entry['Submission_Status']) {
          familyCompletionStatus[familyID] = true;
        }
      });

      // Calculate total and completed families
      totalFamilies = Object.keys(familyCompletionStatus).length;
      completedFamilies = Object.values(familyCompletionStatus).filter(
        (status) => status
      ).length;

      // Calculate the overall participation rate
      const participationRate = (
        (completedFamilies / totalFamilies) *
        100
      ).toFixed(2);

      // Update the submission statuses based on family completion
      data.forEach((entry) => {
        const familyID = entry['Student.Family_ID'];
        if (familyCompletionStatus[familyID]) {
          entry['Submission_Status'] = 1; // Mark as completed
        }
        entry['Participation_Rate'] = participationRate; // Add the participation rate
      });
    } else if (table.include) {
      data = await table.model.findAll({
        include: table.include.map((model) => ({
          model,
          attributes: ['Family_ID'],
        })),
        raw: true,
      });
    } else {
      data = await table.model.findAll({ raw: true });
    }

    // Adding new columns for hyperlinks in the data
    if (table.name === 'Students') {
      data.forEach((row) => {
        row.Family_Reference = ''; // Placeholder for Family VLOOKUP
        row.Bus_Reference = ''; // Placeholder for Bus VLOOKUP
      });
    } else if (
      table.name === 'Survey_Submissions' ||
      table.name === 'Contacts'
    ) {
      data.forEach((row) => {
        row.Family_Reference = ''; // Placeholder for Family VLOOKUP
        row.Student_Reference = ''; // Placeholder for Student VLOOKUP
      });
    } else if (table.name === 'Emergency_Contacts') {
      data.forEach((row) => {
        row.Student_Reference = ''; // Placeholder for Student VLOOKUP
      });
    }

    const sheet = xlsx.utils.json_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, sheet, table.name);
  }

  // Add VLOOKUP formulas for Students sheet
  const studentsSheet = workbook.Sheets['Students'];
  for (let i = 2; i <= 101; i++) {
    studentsSheet[`Y${i}`] = {
      f: `HYPERLINK("#Families!A" & MATCH(Q${i}, Families!A:A, 0), VLOOKUP(Q${i}, Families!A:B, 2, FALSE))`,
    };

    studentsSheet[`Z${i}`] = {
      f: `HYPERLINK("#Buses!A" & MATCH(R${i}, Buses!A:A, 0), VLOOKUP(R${i}, Buses!A:B, 2, FALSE))`,
    };
  }

  // Add VLOOKUP formulas for Survey_Submissions and Contacts sheets
  const surveySubmissionsSheet = workbook.Sheets['Survey_Submissions'];
  for (let i = 2; i <= 101; i++) {
    surveySubmissionsSheet[`K${i}`] = {
      f: `HYPERLINK("#Families!A" & MATCH(J${i}, Families!A:A, 0), VLOOKUP(B${i}, Students!A:Y, 25, FALSE))`,
    };
    // Add Student_Reference formula
    surveySubmissionsSheet[`L${i}`] = {
      f: `HYPERLINK("#Students!A" & MATCH(B${i}, Students!A:A, 0), VLOOKUP(B${i}, Students!A:Y, 2, FALSE))`,
    };
  }

  const contactsSheet = workbook.Sheets['Contacts'];
  for (let i = 2; i <= 51; i++) {
    contactsSheet[`H${i}`] = {
      f: `HYPERLINK("#Families!A" & MATCH(B${i}, Families!A:A, 0), VLOOKUP(B${i}, Families!A:B, 2, FALSE))`,
    };
  }

  // Add VLOOKUP formulas for Emergency_Contacts sheet
  const emergencyContactsSheet = workbook.Sheets['Emergency_Contacts'];
  for (let i = 2; i <= 101; i++) {
    emergencyContactsSheet[`H${i}`] = {
      f: `HYPERLINK("#Students!A" & MATCH(B${i}, Students!A:A, 0), VLOOKUP(B${i}, Students!A:B, 2, FALSE))`,
    };
  }
  // Add Participation_Rate column to Survey_Submissions sheet
  surveySubmissionsSheet['M1'] = { t: 's', v: 'Participation_Rate' };
  for (let i = 2; i <= 51; i++) {
    surveySubmissionsSheet[`M${i}`] = {
      t: 'n',
      v: parseFloat(((completedFamilies / totalFamilies) * 100).toFixed(2)),
    };
  }

  xlsx.writeFile(workbook, 'SchoolData.xlsx');
  console.log('SchoolData.xlsx file has been written with multiple sheets.');
}

// Sync the database and create tables
sequelize.sync({ force: true }).then(async () => {
  console.log('Database & tables created!');

  // Insert sample data
  await createSampleData();
  console.log('Sample data inserted!');

  // Export tables to Excel
  await exportToExcel();
});
