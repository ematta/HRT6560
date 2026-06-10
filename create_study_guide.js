const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        LevelFormat, PageBreak } = require('docx');

// Define custom styles
const doc = new Document({
  styles: {
    default: { 
      document: { 
        run: { font: "Arial", size: 24 } 
      } 
    },
    paragraphStyles: [
      { 
        id: "Heading1", 
        name: "Heading 1", 
        basedOn: "Normal", 
        next: "Normal", 
        quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1F4E79" },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 }
      },
      { 
        id: "Heading2", 
        name: "Heading 2", 
        basedOn: "Normal", 
        next: "Normal", 
        quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 300, after: 150 }, outlineLevel: 1 }
      },
      { 
        id: "Heading3", 
        name: "Heading 3", 
        basedOn: "Normal", 
        next: "Normal", 
        quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "5B9BD5" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 }
      }
    ]
  },
  numbering: {
    config: [
      { 
        reference: "bullets",
        levels: [
          { 
            level: 0, 
            format: LevelFormat.BULLET, 
            text: "•", 
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
      },
      { 
        reference: "numbers",
        levels: [
          { 
            level: 0, 
            format: LevelFormat.DECIMAL, 
            text: "%1.", 
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: {
          width: 12240,
          height: 15840
        },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      // Title Page
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { before: 2400, after: 400 },
        children: [new TextRun("Study Guide")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun("Designing Exceptional Organizational Cultures")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [new TextRun("How to Develop Companies Where Employees Thrive")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 100 },
        children: [new TextRun("by Jamie Jacobs & Hema Crockett")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [new TextRun("Kogan Page, 2021")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [new TextRun("ISBN: 9781789667219")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Table of Contents
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Table of Contents")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Introduction: C is for Conscious")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 1: What is Organizational Culture?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 2: Why Top-Performing Cultures Don't Happen by Accident")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 3: Organizational Structure")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 4: Identifying and Defining Core Values")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 5: Skills for Now and the Future")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 6: Attracting and Retaining the Right Talent")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 7: Engagement and Motivation")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 8: Total Rewards")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 9: Diversity, Inclusion, and Belonging")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Chapter 10: Moving to Action")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Introduction
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Introduction: C is for Conscious")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This book is a practical guide for HR and OD professionals on how to proactively design, build, and foster organizational culture that creates employee and business success. The authors emphasize that exceptional cultures are consciously designed, not left to chance.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Themes")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Culture must be intentionally designed, not accidental")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Employee success drives business success")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Practical, actionable frameworks for HR professionals")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 1
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 1: What is Organizational Culture?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter defines organizational culture and explores the challenges organizations face in understanding and shaping their culture. It examines common pitfalls and provides practical guidance for assessing cultural elements.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Organizational Culture: ", bold: true }), new TextRun("The shared values, beliefs, attitudes, and behaviors that characterize an organization")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Culture as a System: ", bold: true }), new TextRun("Culture encompasses policies, processes, and behaviors")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Cultural Artifacts: ", bold: true }), new TextRun("Visible manifestations of culture (physical environment, language, rituals)")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Assuming culture is just about perks and benefits")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Failing to align culture with business strategy")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Copying another company's culture without adapting to your context")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How would you define your organization's current culture in three words?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What are the visible artifacts of your culture, and what do they communicate?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Where do you see gaps between your stated culture and actual behaviors?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Conduct a cultural audit using the questions provided in the book")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Document visible cultural artifacts and their messages")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Identify three areas where culture and strategy are misaligned")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 2
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 2: Why Top-Performing Cultures Don't Happen by Accident")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter explores how culture operates as a feedback loop and the critical role of leadership in shaping culture. It examines the lifecycle of both toxic and high-performing cultures, emphasizing leader self-awareness.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Culture as Feedback Loop: ", bold: true }), new TextRun("Culture reinforces itself through repeated behaviors and systems")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Leadership Impact: ", bold: true }), new TextRun("Leaders' behaviors have outsized influence on organizational culture")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Self-Awareness: ", bold: true }), new TextRun("Leaders must understand their own impact on culture")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Toxic Culture Lifecycle: ", bold: true }), new TextRun("How lack of self-awareness breeds dysfunction")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "High-Performing Culture Lifecycle: ", bold: true }), new TextRun("How self-aware leaders create positive feedback loops")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Leaders who lack self-awareness about their cultural impact")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Assuming culture will improve without intentional intervention")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Not establishing baseline metrics for cultural assessment")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How does your leadership team model the culture you want to see?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What feedback loops exist in your organization, and are they positive or negative?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How can you increase leadership self-awareness about cultural impact?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Establish baseline culture metrics")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Implement 360-degree feedback for leadership team")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Create a plan for addressing toxic cultural patterns")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 3
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 3: Organizational Structure")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter examines how organizational structure impacts culture, exploring different structural models and their cultural implications. It addresses how structure affects teamwork, leadership expectations, and burnout.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Types of Structures: ", bold: true }), new TextRun("Hierarchical, flat, matrix, networked, and team-based structures")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Structure-Culture Alignment: ", bold: true }), new TextRun("How organizational design supports or hinders desired culture")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Team Dynamics: ", bold: true }), new TextRun("How structure affects collaboration and communication")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Burnout Prevention: ", bold: true }), new TextRun("Structural factors that contribute to or prevent burnout")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Maintaining structures that no longer serve the business strategy")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Unclear roles and responsibilities leading to confusion")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Structures that create silos and hinder collaboration")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Does your current structure support or hinder your desired culture?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How do structural decisions impact employee experience?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What structural changes could reduce burnout in your organization?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Map your current structure against cultural goals")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Identify structural barriers to collaboration")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Develop plan for restructuring if needed")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 4
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 4: Identifying and Defining Core Values")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter provides a framework for identifying, defining, and embedding core values throughout the organization. It emphasizes that values must be operationalized into behaviors and systems.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Core Values Definition: ", bold: true }), new TextRun("Fundamental beliefs that guide organizational behavior and decision-making")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Values vs. Behaviors: ", bold: true }), new TextRun("Values are aspirations; behaviors are how values manifest in action")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Embedding Values: ", bold: true }), new TextRun("Integrating values into hiring, performance management, and daily operations")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Values Assessment: ", bold: true }), new TextRun("Regularly checking if values remain relevant and authentic")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Creating values that are too generic or indistinguishable from competitors")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Not operationalizing values into specific behaviors")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Failing to hold people accountable to values")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Are your stated values actually reflected in daily decisions?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What behaviors would you see if everyone lived your values?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How do you handle situations where values conflict with business results?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Audit current values for authenticity and distinctiveness")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Define 3-5 specific behaviors for each value")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Integrate values into hiring and performance management processes")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 5
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 5: Skills for Now and the Future")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter focuses on identifying and developing the competencies and skills needed for current and future organizational success. It addresses leadership development, competency frameworks, and preparing for future challenges.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Core Competencies: ", bold: true }), new TextRun("Essential skills required across the organization")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Competency Frameworks: ", bold: true }), new TextRun("Structured approaches to defining and developing skills")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Future Skills: ", bold: true }), new TextRun("Anticipating capabilities needed for upcoming challenges")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Leadership Development: ", bold: true }), new TextRun("Building leadership capabilities at all levels")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Assuming leadership is innate rather than developed")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Creating competency frameworks that are too complex or not used")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Focusing only on current skills without planning for the future")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What skills will your organization need in 3-5 years?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How do you currently identify and develop leadership potential?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Are your competency frameworks actively used or sitting on a shelf?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Conduct skills gap analysis for current and future needs")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Simplify competency framework to make it usable")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Create development plans for high-potential employees")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 6
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 6: Attracting and Retaining the Right Talent")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter covers employer branding, employee value propositions (EVP), and recruitment strategies that align with organizational culture. It emphasizes attracting talent that fits values rather than just skills.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Employee Value Proposition (EVP): ", bold: true }), new TextRun("What the organization offers in exchange for employee commitment")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Employer Branding: ", bold: true }), new TextRun("How the organization is perceived as a place to work")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Values Fit vs. Culture Fit: ", bold: true }), new TextRun("Hiring for alignment with core values rather than personality similarity")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Non-Traditional Hires: ", bold: true }), new TextRun("Expanding talent pools beyond conventional criteria")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Hiring for culture fit that creates homogeneity")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Focusing only on technical skills without considering values alignment")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Not being clear about what the organization truly offers employees")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What is your organization's authentic EVP, and is it being delivered?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How do you differentiate between values fit and culture fit in hiring?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What non-traditional talent sources could benefit your organization?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Audit and refine your EVP to ensure authenticity")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Develop values-based interview questions")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Expand recruitment channels to reach diverse talent pools")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 7
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 7: Engagement and Motivation")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter examines strategies for measuring and improving employee engagement and motivation. It covers surveys, feedback mechanisms, and creating meaningful work connections.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Engagement Surveys: ", bold: true }), new TextRun("Tools for measuring employee engagement levels")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Continuous Feedback: ", bold: true }), new TextRun("Beyond annual surveys to ongoing listening")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Skill-Job Alignment: ", bold: true }), new TextRun("Connecting individual strengths to organizational needs")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Motivation Drivers: ", bold: true }), new TextRun("Understanding what truly motivates different individuals")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Conducting surveys without acting on results")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Assuming all employees are motivated by the same things")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Focusing on engagement metrics without understanding underlying issues")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How do you currently measure engagement, and what do you do with the data?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What feedback mechanisms exist beyond formal surveys?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How well are individual skills aligned with role requirements?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Review current engagement survey process and action planning")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Implement regular check-ins between managers and employees")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Create skill-job alignment analysis for key roles")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 8
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 8: Total Rewards")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter explores comprehensive reward strategies that go beyond compensation to include benefits, recognition, and non-monetary rewards. It emphasizes personalization and alignment with cultural values.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Total Rewards Components: ", bold: true }), new TextRun("Compensation, benefits, recognition, development, and work environment")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Personalization: ", bold: true }), new TextRun("Tailoring rewards to individual preferences and life stages")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Non-Monetary Rewards: ", bold: true }), new TextRun("Recognition, flexibility, development opportunities, and meaningful work")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Reward Philosophy: ", bold: true }), new TextRun("How rewards reflect and reinforce organizational values")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Focusing only on compensation while ignoring other reward elements")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Offering one-size-fits-all rewards that don't meet diverse needs")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Not communicating the full value of the rewards package")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What do your current rewards say about your organizational values?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How well do your rewards meet the diverse needs of your workforce?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What non-monetary rewards could have the biggest impact?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Audit total rewards package for alignment with values")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Survey employees on reward preferences")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Develop communication strategy for total rewards value")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 9
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 9: Diversity, Inclusion, and Belonging")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This chapter addresses creating an inclusive culture where diverse perspectives are valued. It covers moving beyond compliance to genuine inclusion, measuring progress, and the business case for diversity.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Diversity vs. Inclusion vs. Belonging: ", bold: true }), new TextRun("Representation is not enough; people must feel valued and included")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Beyond Quotas: ", bold: true }), new TextRun("Moving from compliance to genuine cultural integration")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Measuring ROI: ", bold: true }), new TextRun("Quantifying the business impact of diversity initiatives")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Leadership Accountability: ", bold: true }), new TextRun("Holding leaders responsible for inclusive behaviors")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Treating diversity as a numbers game without addressing inclusion")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Not addressing unconscious bias in systems and processes")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Failing to hold leadership accountable for inclusive behaviors")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How do you know if employees from underrepresented groups feel they belong?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What systems or processes might inadvertently create barriers to inclusion?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How do you measure the ROI of your diversity and inclusion efforts?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Conduct inclusion assessment beyond demographic data")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Audit hiring and promotion processes for bias")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Set leadership accountability metrics for inclusion")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Chapter 10
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Chapter 10: Moving to Action")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Chapter Summary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This final chapter brings together all elements into an actionable change management plan. It provides frameworks for implementation, tracking progress, and sustaining cultural transformation.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Key Concepts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Change Management: ", bold: true }), new TextRun("Structured approach to cultural transformation")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Implementation Planning: ", bold: true }), new TextRun("Prioritizing and sequencing cultural initiatives")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Sustainability: ", bold: true }), new TextRun("Ensuring cultural changes endure over time")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun({ text: "Metrics and Accountability: ", bold: true }), new TextRun("Tracking progress and holding people accountable")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Common Pitfalls")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Trying to change everything at once instead of prioritizing")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Not involving employees in the change process")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Lack of sustained focus after initial launch")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Discussion Questions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What is your highest priority cultural initiative?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How will you sequence cultural changes for maximum impact?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What mechanisms will ensure sustained focus on culture?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Action Items")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Create prioritized action plan with timelines")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Establish governance structure for culture initiatives")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Define metrics and review cadence")]
      }),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Key Terms Glossary
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Key Terms Glossary")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This glossary defines important terms used throughout the book.")]
      }),
      createGlossaryEntry("Core Values", "Fundamental beliefs that guide organizational behavior and decision-making."),
      createGlossaryEntry("Culture Fit", "Alignment between an individual's values and behaviors and the organization's culture."),
      createGlossaryEntry("Employee Engagement", "The emotional commitment employees have to the organization and its goals."),
      createGlossaryEntry("Employee Value Proposition (EVP)", "The unique set of benefits and values an organization offers in exchange for employee skills and commitment."),
      createGlossaryEntry("Organizational Culture", "The shared values, beliefs, attitudes, and behaviors that characterize how work gets done in an organization."),
      createGlossaryEntry("Self-Awareness", "Understanding one's own personality, behaviors, and impact on others."),
      createGlossaryEntry("Total Rewards", "The complete package of compensation, benefits, recognition, and developmental opportunities provided to employees."),
      createGlossaryEntry("Values Fit", "Alignment between an individual's personal values and the organization's stated core values."),
      
      new Paragraph({ children: [new PageBreak()] }),
      
      // Final Reflection
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Final Reflection Questions")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("Use these questions to reflect on the book as a whole and plan your cultural transformation journey.")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Overall Assessment")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("On a scale of 1-10, how would you rate your current organizational culture?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What are the three biggest strengths of your current culture?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What are the three biggest gaps or weaknesses?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Priority Actions")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What is the one cultural change that would have the biggest impact?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Who needs to be involved to make this change successful?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What resources (time, money, expertise) do you need?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What is your timeline for implementation?")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Sustainability Plan")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How will you ensure cultural changes stick over time?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("What metrics will you track to measure progress?")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("How often will you review and adjust your cultural strategy?")]
      }),
    ]
  }]
});

// Helper function to create glossary entries
function createGlossaryEntry(term, definition) {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    children: [
      new TextRun({ text: term + ": ", bold: true }),
      new TextRun(definition)
    ]
  });
}

// Generate the document
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Designing_Exceptional_Organizational_Cultures_Study_Guide.docx", buffer);
  console.log("Study guide created successfully!");
}).catch(err => {
  console.error("Error creating document:", err);
});
