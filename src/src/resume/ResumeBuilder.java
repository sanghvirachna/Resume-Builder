package src.resume;
import src.resume.word;
import javax.swing.*;
import javax.swing.border.LineBorder;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;


public class ResumeBuilder extends JFrame {

    private JPanel mainPanel;
    private CardLayout cardLayout;
    private JPanel generalInfoPanel;
    private JPanel educationPanel;
    private JPanel projectsPanel;
    private JPanel experiencePanel;
    
    
    public ResumeBuilder() {
        setTitle("Create Your Resume");
        setSize(900, 800);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        mainPanel = new JPanel();
        cardLayout = new CardLayout();
        mainPanel.setLayout(cardLayout);
        mainPanel.setBackground(Color.BLACK);

        generalInfoPanel = createGeneralInfoPanel();
        educationPanel = createEducationPanel();
        projectsPanel = createProjectsPanel();
        experiencePanel = createExperiencePanel();

        mainPanel.add(generalInfoPanel, "General Info");
        mainPanel.add(educationPanel, "Education");
        mainPanel.add(projectsPanel, "Projects");
        mainPanel.add(experiencePanel, "Experience");


        JButton generalInfoButton = new JButton("General Info");
        generalInfoButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                cardLayout.show(mainPanel, "General Info");
            }
        });

        JButton educationButton = new JButton("Education");
        educationButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                cardLayout.show(mainPanel, "Education");
            }
        });

        JButton projectsButton = new JButton("Projects");
        projectsButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                cardLayout.show(mainPanel, "Projects");
            }
        });

        JButton experienceButton = new JButton("Experience");
        experienceButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                cardLayout.show(mainPanel, "Experience");
            }
        });
        

        JPanel buttonPanel = new JPanel();
        buttonPanel.add(generalInfoButton);
        buttonPanel.add(educationButton);
        buttonPanel.add(projectsButton);
        buttonPanel.add(experienceButton);

        add(buttonPanel, BorderLayout.NORTH);
        add(mainPanel, BorderLayout.CENTER);

        setVisible(true);
    }
    String name;
	String email;
	String linkedIn;
	String github;
	String contact;
	String skill;
	String objective;
    
    private JPanel createGeneralInfoPanel() {
    	JPanel panel = new JPanel();
    	panel.setLayout(new GridBagLayout());
    	panel.setBackground(new Color(35,43,42));
    	GridBagConstraints constraints = new GridBagConstraints();
    	constraints.fill = GridBagConstraints.HORIZONTAL;
    	constraints.insets = new Insets(10, 10, 10, 10);

    	JLabel nameLabel = new JLabel("Name:");
    	nameLabel.setForeground(Color.WHITE);
    	JTextField nameTextField = new JTextField();
    	nameTextField.setBackground(Color.BLACK);
    	nameTextField.setForeground(Color.WHITE);
    	nameTextField.setBorder(new LineBorder(Color.WHITE));
    	nameLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18
    	nameTextField.setFont(new Font("Arial", Font.PLAIN, 18)); // Set font size to 18
    	nameTextField.setColumns(20);



    	JLabel emailLabel = new JLabel("Email:");
    	emailLabel.setForeground(Color.WHITE);
    	JTextField emailTextField = new JTextField();
    	emailTextField.setBackground(Color.BLACK);
    	emailTextField.setForeground(Color.WHITE);
    	emailTextField.setBorder(new LineBorder(Color.WHITE));
    	emailLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18
    	emailTextField.setFont(new Font("Arial", Font.PLAIN, 18)); // Set font size to 18



    	JLabel contactNoLabel = new JLabel("Contact No:");
    	contactNoLabel.setForeground(Color.WHITE);
    	JTextField contactNoTextField = new JTextField();
    	contactNoTextField.setBackground(Color.BLACK);
    	contactNoTextField.setForeground(Color.WHITE);
    	contactNoTextField.setBorder(new LineBorder(Color.WHITE));
    	contactNoLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18
    	contactNoTextField.setFont(new Font("Arial", Font.PLAIN, 18)); // Set font size to 18



    	JLabel linkedInLabel = new JLabel("LinkedIn Profile:");
    	linkedInLabel.setForeground(Color.WHITE);
    	JTextField linkedInTextField = new JTextField();
    	linkedInTextField.setBackground(Color.BLACK);
    	linkedInTextField.setForeground(Color.WHITE);
    	linkedInTextField.setBorder(new LineBorder(Color.WHITE));
    	linkedInLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18
    	linkedInTextField.setFont(new Font("Arial", Font.PLAIN, 18)); // Set font size to 18
    	
    	JLabel skillLabel = new JLabel("Skills : ");
    	skillLabel.setForeground(Color.WHITE);
    	
    	JTextArea skillTextField = new JTextArea(10,5);
    	skillTextField.setBackground(Color.BLACK);
    	skillTextField.setForeground(Color.WHITE);
    	skillTextField.setBorder(new LineBorder(Color.WHITE));
    	skillLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18
    	skillTextField.setFont(new Font("Arial", Font.PLAIN, 16)); // Set font size to 18

    	
    	JLabel objectiveLabel = new JLabel("Objective : ");
    	objectiveLabel.setForeground(Color.WHITE);
    	
    	JTextArea objectiveTextField = new JTextArea(10,5);
    	objectiveTextField.setBackground(Color.BLACK);
    	objectiveTextField.setForeground(Color.WHITE);
    	objectiveTextField.setBorder(new LineBorder(Color.WHITE));
    	objectiveLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18
    	objectiveTextField.setFont(new Font("Arial", Font.PLAIN, 16)); // Set font size to 18




    	JLabel githubLabel = new JLabel("GitHub Profile:");
    	githubLabel.setForeground(Color.WHITE);
    	JTextField githubTextField = new JTextField();
    	githubTextField.setBackground(Color.BLACK);
    	githubTextField.setForeground(Color.WHITE);
    	githubTextField.setBorder(new LineBorder(Color.WHITE));
    	githubLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18
    	githubTextField.setFont(new Font("Arial", Font.PLAIN, 18)); // Set font size to 18

    	
    	
        JButton nextButton = new JButton("Next");
        nextButton.setBackground(new Color(17,122,101));
        nextButton.setForeground(Color.WHITE);
        nextButton.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

       

    	
    	
 	  nextButton.addActionListener(new ActionListener() {
          @Override
          public void actionPerformed(ActionEvent e) {
        	  name = nameTextField.getText();
       	    github = githubTextField.getText();
       	    contact = contactNoTextField.getText();
       	   email = emailTextField.getText();
       	   linkedIn = linkedInTextField.getText();
       	   skill = skillTextField.getText();
       	   objective = objectiveTextField.getText();
       	   

              System.out.println(name+github+contact+email+linkedIn);
              createEducationPanel();
              cardLayout.show(mainPanel, "Education");
          }
      });
      
    	constraints.gridx = 0;
    	constraints.gridy = 0;
    	panel.add(nameLabel, constraints);

    	constraints.gridx = 1;
    	constraints.gridy = 0;
    	constraints.gridwidth = 2;
    	panel.add(nameTextField, constraints);

    	constraints.gridx = 0;
    	constraints.gridy = 1;
    	constraints.gridwidth = 1;
    	panel.add(emailLabel, constraints);

    	constraints.gridx = 1;
    	constraints.gridy = 1;
    	constraints.gridwidth = 2;
    	panel.add(emailTextField, constraints);

    	constraints.gridx = 0;
    	constraints.gridy = 2;
    	constraints.gridwidth = 1;
    	panel.add(contactNoLabel, constraints);

    	constraints.gridx = 1;
    	constraints.gridy = 2;
    	constraints.gridwidth = 2;
    	panel.add(contactNoTextField, constraints);

    	constraints.gridx = 0;
    	constraints.gridy = 3;
    	constraints.gridwidth = 1;
    	panel.add(linkedInLabel, constraints);

    	constraints.gridx = 1;
    	constraints.gridy = 3;
    	constraints.gridwidth = 2;
    	panel.add(linkedInTextField, constraints);

    	constraints.gridx = 0;
    	constraints.gridy = 4;
    	constraints.gridwidth = 1;
    	panel.add(githubLabel, constraints);

    	constraints.gridx = 1;
    	constraints.gridy = 4;
    	constraints.gridwidth = 2;
    	panel.add(githubTextField, constraints);
    	
    	constraints.gridx = 0;
        constraints.gridy = 5;
        constraints.gridwidth = 1;
        constraints.anchor = GridBagConstraints.LINE_END;
        panel.add(skillLabel, constraints);

        constraints.gridx = 1;
        constraints.gridy = 5;
        constraints.gridwidth = 2;
        constraints.anchor = GridBagConstraints.LINE_START;
        panel.add(skillTextField, constraints);   
        
        constraints.gridx = 0;
        constraints.gridy = 6;
        constraints.gridwidth = 1;
        constraints.anchor = GridBagConstraints.LINE_END;
        panel.add(objectiveLabel, constraints);

        constraints.gridx = 1;
        constraints.gridy = 6;
        constraints.gridwidth = 2;
        constraints.anchor = GridBagConstraints.LINE_START;
        panel.add(objectiveTextField, constraints);    	

        constraints.gridx = 1;
    	constraints.gridy = 5;
    	constraints.gridwidth = 1;
    	constraints.anchor = GridBagConstraints.CENTER;
    	panel.add(nextButton, constraints);

    	
    	constraints.gridx = 0;
    	constraints.gridy = 7;
    	constraints.gridwidth=7;
    	panel.add(nextButton, constraints);

    	return panel;
  

        }
           
    private String[][] subsectionsData = new String[4][4]; // Array to store subsection data

    private JPanel createEducationPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BorderLayout());
        panel.setBackground(new Color(35,43,42));

        JPanel subsectionsPanel = new JPanel(new GridLayout(0, 1));
        subsectionsPanel.setBackground(new Color(35,43,42));

        // Create four subsections
        for (int i = 0; i < 4; i++) {
            JPanel subsection = createSubsectionPanel(i);
            subsectionsPanel.add(subsection);
        }

        JScrollPane scrollPane = new JScrollPane(subsectionsPanel);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        scrollPane.getViewport().setBackground(new Color(35,43,42));

        JButton nextButton = new JButton("Next");
        nextButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Save the entered education data and proceed to the Projects section
                saveEducationData();
                for(int i = 0 ; i < 4 ; i ++) {
                	for(int j = 0 ;  j < 4 ; j++) {
                		System.out.print(subsectionsData[i][j]);
                	}
                }
                cardLayout.show(mainPanel,"Projects");
                
            }
        });
        
        
        nextButton.setBackground(new Color(17,122,101));
        nextButton.setForeground(Color.WHITE);
        nextButton.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18


        panel.add(scrollPane, BorderLayout.CENTER);
        panel.add(nextButton, BorderLayout.SOUTH);

        return panel;
    }

    private JPanel createSubsectionPanel(int index) {
        JPanel subsection = new JPanel();
        subsection.setLayout(new GridBagLayout());
        subsection.setBackground(new Color(35,43,42));

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.CENTER; // Center alignment
        gbc.insets = new Insets(5, 10, 5, 10);

        JLabel headingLabel = new JLabel("Education " + (index + 1));
        headingLabel.setFont(new Font("Arial", Font.BOLD, 20));
        headingLabel.setForeground(Color.WHITE);

        JLabel educationLabel = new JLabel("Education:");
        educationLabel.setForeground(Color.WHITE);
        educationLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18
    	
        JTextField educationTextField = new JTextField(20);
        educationTextField.setBackground(Color.BLACK);
        educationTextField.setForeground(Color.WHITE);
        educationTextField.setBorder(new LineBorder(Color.WHITE));

        JLabel schoolLabel = new JLabel("School/University:");
        schoolLabel.setForeground(Color.WHITE);
        schoolLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        JTextField schoolTextField = new JTextField(20);
        schoolTextField.setBackground(Color.BLACK);
        schoolTextField.setForeground(Color.WHITE);
        schoolTextField.setBorder(new LineBorder(Color.WHITE));

        JLabel yearLabel = new JLabel("Year:");
        yearLabel.setForeground(Color.WHITE);
        yearLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        JTextField yearTextField = new JTextField(20);
        yearTextField.setBackground(Color.BLACK);
        yearTextField.setForeground(Color.WHITE);
        yearTextField.setBorder(new LineBorder(Color.WHITE));

        JLabel cityLabel = new JLabel("City/State:");
        cityLabel.setForeground(Color.WHITE);
        cityLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        
        JTextField cityTextField = new JTextField(20);
        cityTextField.setBackground(Color.BLACK);
        cityTextField.setForeground(Color.WHITE);
        cityTextField.setBorder(new LineBorder(Color.WHITE));
        
        

        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2; // Span across 2 columns
        subsection.add(headingLabel, gbc);

        gbc.gridy++;
        gbc.gridwidth = 1; // Reset gridwidth
        gbc.anchor = GridBagConstraints.WEST;
        subsection.add(educationLabel, gbc);

        gbc.gridy++;
        subsection.add(schoolLabel, gbc);

        gbc.gridy++;
        subsection.add(yearLabel, gbc);

        gbc.gridy++;
        subsection.add(cityLabel, gbc);

        gbc.gridx = 1;
        gbc.gridy = 1;
        gbc.anchor = GridBagConstraints.EAST;
        subsection.add(educationTextField, gbc);

        gbc.gridy++;
        subsection.add(schoolTextField, gbc);

        gbc.gridy++;
        subsection.add(yearTextField, gbc);

        gbc.gridy++;
        subsection.add(cityTextField, gbc);

        gbc.gridx = 0;
        gbc.gridy++;
        gbc.gridwidth = 2;
        gbc.anchor = GridBagConstraints.CENTER;
        gbc.insets = new Insets(10, 10, 10, 10);

        JButton addButton = new JButton("Add");
        addButton.setBackground(new Color(0, 122, 204)); // Light blue color
        addButton.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        addButton.setForeground(Color.WHITE);
        subsection.add(addButton, gbc);

        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Add the education subsection to the subsectionsData array
                subsectionsData[index][0] = educationTextField.getText();
                subsectionsData[index][1] = schoolTextField.getText();
                subsectionsData[index][2] = yearTextField.getText();
                subsectionsData[index][3] = cityTextField.getText();
                System.out.println("Education " + (index + 1) + " added.");
            }
        });

        return subsection;
    }
    private void saveEducationData() {
        // Display the saved education data
        for (int i = 0; i < 4; i++) {
            String education = subsectionsData[i][0] != null ? subsectionsData[i][0] : "";
            String school = subsectionsData[i][1] != null ? subsectionsData[i][1] : "";
            String year = subsectionsData[i][2] != null ? subsectionsData[i][2] : "";
            String city = subsectionsData[i][3] != null ? subsectionsData[i][3] : "";

//            System.out.println("Education " + (i + 1));
//            System.out.println("Education: " + education);
//            System.out.println("School/University: " + school);
//            System.out.println("Year: " + year);
//            System.out.println("City/State: " + city);
//            System.out.println("-------------------------");
        }
    }

    

    private String[][] subsectionsProjectData = new String[4][2]; // Array to store subsection data

    private JPanel createProjectsPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BorderLayout());
        panel.setBackground(new Color(35,43,42));

        JPanel subsectionsPanel = new JPanel(new GridLayout(0, 1));
        subsectionsPanel.setBackground(new Color(35,43,42));

        // Create four subsections
        for (int i = 0; i < 4; i++) {
            JPanel subsection = createSubsectionProjectPanel(i);
            subsectionsPanel.add(subsection);
        }

        JScrollPane scrollPane = new JScrollPane(subsectionsPanel);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        scrollPane.getViewport().setBackground(new Color(35,43,42));

        JButton nextButton = new JButton("Next");
        
        nextButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Save the entered education data and proceed to the Projects section
                saveProjectData();
                cardLayout.show(mainPanel, "Experience");

                
            }
        });
        nextButton.setBackground(new Color(17, 122, 101));
        nextButton.setForeground(Color.WHITE);
        nextButton.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18


        panel.add(scrollPane, BorderLayout.CENTER);
        panel.add(nextButton, BorderLayout.SOUTH);

        return panel;
    }

    private JPanel createSubsectionProjectPanel(int index) {
        JPanel subsection = new JPanel();
        subsection.setLayout(new GridBagLayout());
        subsection.setBackground(new Color(35,43,42));

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.CENTER; // Center alignment
        gbc.insets = new Insets(5, 10, 5, 10);

        JLabel headingLabel = new JLabel("Project " + (index + 1));
        headingLabel.setFont(new Font("Arial", Font.BOLD, 20));
        headingLabel.setForeground(Color.WHITE);

        JLabel titleLabel = new JLabel("Title:");
        titleLabel.setForeground(Color.WHITE);
        titleLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        JTextArea educationTextField = new JTextArea(4,30);
        
        educationTextField.setBackground(Color.BLACK);
        educationTextField.setForeground(Color.WHITE);
        educationTextField.setBorder(new LineBorder(Color.WHITE));

        JLabel descriptionLabel = new JLabel("Description:");
        descriptionLabel.setForeground(Color.WHITE);
        descriptionLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        JTextArea schoolTextField = new JTextArea(8,30);
        schoolTextField.setBackground(Color.BLACK);
        schoolTextField.setForeground(Color.WHITE);
        schoolTextField.setBorder(new LineBorder(Color.WHITE));

        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2; // Span across 2 columns
        subsection.add(headingLabel, gbc);

        gbc.gridy++;
        gbc.gridwidth = 1; // Reset gridwidth
        gbc.anchor = GridBagConstraints.WEST;
        subsection.add(titleLabel, gbc);

        gbc.gridy++;
        subsection.add(descriptionLabel, gbc);

        
        gbc.gridx = 1;
        gbc.gridy = 1;
        gbc.anchor = GridBagConstraints.EAST;
        subsection.add(educationTextField, gbc);

        gbc.gridy++;
        subsection.add(schoolTextField, gbc);

        
        gbc.gridx = 0;
        gbc.gridy++;
        gbc.gridwidth = 2;
        gbc.anchor = GridBagConstraints.CENTER;
        gbc.insets = new Insets(10, 10, 10, 10);

        JButton addButton = new JButton("Add");
        addButton.setBackground(new Color(0, 122, 204)); // Light blue color
        addButton.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        addButton.setForeground(Color.WHITE);
        subsection.add(addButton, gbc);

        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Add the education subsection to the subsectionsData array
                subsectionsProjectData[index][0] = educationTextField.getText();
                subsectionsProjectData[index][1] = schoolTextField.getText();
                System.out.println("Education " + (index + 1) + " added.");
            }
        });

        return subsection;
    }
    private void saveProjectData() {
        // Display the saved education data
        for (int i = 0; i < 4; i++) {
            String title = subsectionsProjectData[i][0] != null ? subsectionsProjectData[i][0] : "";
            String description = subsectionsProjectData[i][1] != null ? subsectionsProjectData[i][1] : "";
            
            System.out.println("Title " + (i + 1) + title);
            System.out.println("Description: " + description);
            
        }
    }

    private String[][] subsectionsExperienceData = new String[4][3]; // Array to store subsection data

    private JPanel createExperiencePanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BorderLayout());
        panel.setBackground(new Color(35,43,42));

        JPanel subsectionsPanel = new JPanel(new GridLayout(0, 1));
        subsectionsPanel.setBackground(new Color(35,43,42));

        // Create four subsections
        for (int i = 0; i < 4; i++) {
            JPanel subsection = createSubsectionExperiencePanel(i);
            subsectionsPanel.add(subsection);
        }

        JScrollPane scrollPane = new JScrollPane(subsectionsPanel);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        scrollPane.getViewport().setBackground(new Color(35,43,42));

        JButton nextButton = new JButton("Create Resume");
        nextButton.setFont(new Font("Arial", Font.BOLD, 24)); // Set font size to 18

        nextButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Save the entered education data and proceed to the Projects section
                saveExperienceData();
                try {
					word();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
                
            }
        });
        nextButton.setBackground(new Color(17, 122, 101));
        nextButton.setForeground(Color.WHITE);
        nextButton.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18


        panel.add(scrollPane, BorderLayout.CENTER);
        panel.add(nextButton, BorderLayout.SOUTH);

        return panel;
    }

    private JPanel createSubsectionExperiencePanel(int index) {
        JPanel subsection = new JPanel();
        subsection.setLayout(new GridBagLayout());
        subsection.setBackground(new Color(35,43,42));

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.CENTER; // Center alignment
        gbc.insets = new Insets(5, 10, 5, 10);

        JLabel headingLabel = new JLabel("Experience " + (index + 1));
        headingLabel.setFont(new Font("Arial", Font.BOLD, 20));
        headingLabel.setForeground(Color.WHITE);

        JLabel titleLabel = new JLabel("Job Role:");
        titleLabel.setForeground(Color.WHITE);
        JTextArea educationTextField = new JTextArea(3,30);
        titleLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        
        educationTextField.setBackground(Color.BLACK);
        educationTextField.setForeground(Color.WHITE);
        educationTextField.setBorder(new LineBorder(Color.WHITE));

        
       
        
        JLabel descriptionLabel = new JLabel("Description:");
        descriptionLabel.setForeground(Color.WHITE);
        descriptionLabel.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        JTextArea schoolTextField = new JTextArea(8,30);
        schoolTextField.setBackground(Color.BLACK);
        schoolTextField.setForeground(Color.WHITE);
        schoolTextField.setBorder(new LineBorder(Color.WHITE));

        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2; // Span across 2 columns
        subsection.add(headingLabel, gbc);

        gbc.gridy++;
        gbc.gridwidth = 1; // Reset gridwidth
        gbc.anchor = GridBagConstraints.WEST;
        subsection.add(titleLabel, gbc);

        gbc.gridy++;
        subsection.add(descriptionLabel, gbc);

        
        gbc.gridx = 1;
        gbc.gridy = 1;
        gbc.anchor = GridBagConstraints.EAST;
        subsection.add(educationTextField, gbc);

        gbc.gridy++;
        subsection.add(schoolTextField, gbc);

        
        gbc.gridx = 0;
        gbc.gridy++;
        gbc.gridwidth = 2;
        gbc.anchor = GridBagConstraints.CENTER;
        gbc.insets = new Insets(10, 10, 10, 10);

        JButton addButton = new JButton("Add");
        addButton.setFont(new Font("Arial", Font.BOLD, 18)); // Set font size to 18

        addButton.setBackground(new Color(0, 122, 204)); // Light blue color
        addButton.setForeground(Color.WHITE);
        subsection.add(addButton, gbc);

        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Add the education subsection to the subsectionsData array
                subsectionsExperienceData[index][0] = educationTextField.getText();
                subsectionsExperienceData[index][1] = schoolTextField.getText();
                System.out.println("Education " + (index + 1) + " added.");
            }
        });

        return subsection;
    }
    private void saveExperienceData() {
        // Display the saved experience data
        for (int i = 0; i < 4; i++) {
            String JobRole = subsectionsExperienceData[i][0] != null ? subsectionsExperienceData[i][0] : "";
            String period = subsectionsExperienceData[i][1] != null ? subsectionsExperienceData[i][1] : "";
            String description = subsectionsExperienceData[i][0] != null ? subsectionsExperienceData[i][2] : "";

            System.out.println("Job Role " + (i + 1) + JobRole);
            System.out.println("Period " + (i + 1) + period);
            System.out.println("Description: " + description);
            
        }
    }
    private void word() throws IOException {
    	       XWPFDocument document = new XWPFDocument(); 
    	    		
    	          //Write the Document in file system
    	          FileOutputStream out = new FileOutputStream( new File("createdocument.docx"));
    	          String nameValue = name;
    	          
    	        //Create a blank spreadsheet
    	          
    	          XWPFHeaderFooterPolicy headerFooterPolicy = document.getHeaderFooterPolicy();
    	          if (headerFooterPolicy == null) headerFooterPolicy = document.createHeaderFooterPolicy();

    	          // create header start
    	          XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

    	          XWPFParagraph paragraph = header.createParagraph();
    	          paragraph.setAlignment(ParagraphAlignment.LEFT);

    	          XWPFRun run = paragraph.createRun();  
    	          run.setText(nameValue);
    	          run.setFontSize(50);
    	          run.setFontFamily("Bahnschrift Light", null);
    	          run.setColor("000000");
    	          
    	              XWPFParagraph lineBreakParagraph = document.createParagraph();
    	              XWPFRun lineBreakRun = lineBreakParagraph.createRun();
    	          
    	              XWPFTable table = document.createTable(1, 2);
    	              table.setWidth("100%");

    	              CTTblWidth[] colWidths = new CTTblWidth[] {
    	                      table.getRow(0).getCell(0).getCTTc().addNewTcPr().addNewTcW(),
    	                      table.getRow(0).getCell(1).getCTTc().addNewTcPr().addNewTcW()
    	                  };
    	                  colWidths[0].setType(STTblWidth.PCT);
    	                  colWidths[1].setType(STTblWidth.PCT);
    	                  colWidths[0].setW("50%");
    	                  colWidths[1].setW("50%");

    	                  // Get the table row and cells
    	                  XWPFTableRow row = table.getRow(0);
    	                  XWPFTableCell cell1 = row.getCell(0);
    	                  XWPFTableCell cell2 = row.getCell(1);


        	              table.getCTTbl().getTblPr().unsetTblBorders();

    	              // Format cell1 with variables on separate lines
    	              cell1.setVerticalAlignment(XWPFVertAlign.CENTER);

    	              String emailValue = email;
    	              String contactValue = contact;
    	              String linkedinValue = linkedIn;
    	              String githubValue = github;
    	             
    	              XWPFParagraph emaillabel = cell1.getParagraphArray(0);
    	              XWPFRun cell1Run0 = emaillabel.createRun();
    	              cell1Run0.setText("Email :");
    	              cell1Run0.setFontSize(18);
        	          cell1Run0.setFontFamily("Bahnschrift", null);

    	              
    	              XWPFParagraph cell1Paragraph = cell1.getParagraphArray(0);
    	              XWPFRun cell1Run00 = cell1Paragraph.createRun();
    	              cell1Run00.setText(emailValue);
    	              cell1Run00.setFontSize(16);
        	          cell1Run00.setFontFamily("Bahnschrift", null);

    	              
    	              cell1Paragraph.createRun().addBreak();

    	              XWPFParagraph contactlabel = cell1.getParagraphArray(0);
    	              XWPFRun cell1Run11 = contactlabel.createRun();
    	              cell1Run11.setText("Contact :");
    	              cell1Run11.setFontSize(18);
        	          cell1Run11.setFontFamily("Bahnschrift", null);

    	              
    	              XWPFRun cell1Run2 = cell1Paragraph.createRun();
    	              cell1Run2.setText(contactValue);
    	              cell1Run2.setFontSize(16);
        	          cell1Run2.setFontFamily("Bahnschrift", null);


    	              
    	              cell1Paragraph.createRun().addBreak();
    	              
    	              XWPFParagraph linkedinlabel = cell1.getParagraphArray(0);
    	              XWPFRun cell1Run12 = linkedinlabel.createRun();
    	              cell1Run12.setText("LinkedIn :");
    	              cell1Run12.setFontSize(18);
        	          cell1Run12.setFontFamily("Bahnschrift", null);

    	              
    	              XWPFRun cell1Run3 = cell1Paragraph.createRun();
    	              cell1Run3.setText(linkedinValue);
    	              cell1Run3.setFontSize(16);
        	          cell1Run3.setFontFamily("Bahnschrift", null);


    	              
    	              cell1Paragraph.createRun().addBreak();

    	              XWPFParagraph githublabel = cell1.getParagraphArray(0);
    	              XWPFRun cell1Run13 = githublabel.createRun();
    	              cell1Run13.setText("Github :");
    	              cell1Run13.setFontSize(18);
        	          cell1Run13.setFontFamily("Bahnschrift", null);

    	              
    	              XWPFRun cell1Run4 = cell1Paragraph.createRun();
    	              cell1Run4.setText(githubValue);
    	              cell1Run4.setFontSize(16);
        	          cell1Run4.setFontFamily("Bahnschrift", null);



    	              // Format cell2 with objective heading and paragraph
    	              cell2.setVerticalAlignment(XWPFVertAlign.CENTER);

    	              String objective1 = objective;
    	              XWPFParagraph cell2Paragraph = cell2.getParagraphArray(0);
    	              XWPFRun cell2Run1 = cell2Paragraph.createRun();
    	              cell2Run1.setText("Objective".toUpperCase());
        	          cell2Run1.setFontFamily("Bahnschrift", null);

    	              cell2Run1.setFontSize(25);
    	              cell2Paragraph.createRun().addBreak();
    	              XWPFRun cell2Run2 = cell2Paragraph.createRun();
    	              cell2Run2.setText(objective1);
    	              cell2Run2.setFontSize(14);
    	              
//    	              //1 section done 
//    	              
//    	              
    	              XWPFTable table1 = document.createTable();
    	              table1.setWidth("100%");

    	              // Remove all borders from table1
    	              table1.getCTTbl().getTblPr().unsetTblBorders();

    	              // Set column widths
    	              CTTblWidth[] colWidths1 = new CTTblWidth[]{
    	                      table1.getRow(0).addNewTableCell().getCTTc().addNewTcPr().addNewTcW(),
    	                      table1.getRow(0).addNewTableCell().getCTTc().addNewTcPr().addNewTcW()
    	              };
    	              colWidths1[0].setType(STTblWidth.PCT);
    	              colWidths1[1].setType(STTblWidth.PCT);
    	              colWidths1[0].setW("50%");
    	              colWidths1[1].setW("50%");

    	              // Get the table cells
    	              XWPFTableCell cell01 = table1.getRow(0).getCell(0);
    	              XWPFTableCell cell02 = table1.getRow(0).getCell(1);
    	              XWPFTableCell cell3 = table1.createRow().getCell(0);
    	              XWPFTableCell cell4 = table1.getRow(1).getCell(0);

    	              // Format cell1 with "Skills" heading and paragraphs
    	              cell01.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

    	              String skills = skill;

    	              XWPFParagraph cell01Paragraph = cell01.getParagraphArray(0);
    	              XWPFRun cell01Run1 = cell01Paragraph.createRun();
    	              cell01Run1.setText("Skills".toUpperCase());
    	              cell01Run1.setFontSize(25);
    	              cell01Paragraph.createRun().addBreak();
    	              XWPFRun cell01Run2 = cell01Paragraph.createRun();
    	              cell01Run2.setText(skills);

    	              // Format cell2 with "Education" heading and data from 2D array
    	              cell02.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

    	              String[][] educationData = subsectionsData;

    	              XWPFParagraph cell02Paragraph = cell02.getParagraphArray(0);
    	              XWPFRun cell02Run1 = cell02Paragraph.createRun();
    	              cell02Run1.setText("Education".toUpperCase());
    	              cell02Run1.setFontSize(25);

    	              for (String[] data : educationData) {
    	                  cell2Paragraph.createRun().addBreak();
    	                  XWPFRun cell02Run2 = cell02Paragraph.createRun();
    	                  cell02Run2.setText(data[0] + " - " + data[1] + " - " + data[2]);
    	              }

    	              // Format cell3 with "Experience" heading and data from 2D array
    	              cell3.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

    	              String[][] experienceData = subsectionsExperienceData;

    	              XWPFParagraph cell3Paragraph = cell3.getParagraphArray(0);
    	              XWPFRun cell3Run1 = cell3Paragraph.createRun();
    	              cell3Run1.setText("Experience".toUpperCase());
    	              cell3Run1.setFontSize(25);

    	              for (String[] data : experienceData) {
    	                  if (data != null && data.length >= 2 && data[0] != null && data[1] != null) {
    	                      cell3Paragraph.createRun().addBreak();

    	                      // Print Job Role (Bold, Font Size: 18)
    	                      XWPFRun cell3Run2 = cell3Paragraph.createRun();
    	                      cell3Run2.setBold(true);
    	                      cell3Run2.setFontSize(18);
    	                      cell3Run2.setText("Job Role - " + data[0]);

    	                      cell3Paragraph.createRun().addBreak();

    	                      // Print Description (Font Size: 18)
    	                      XWPFRun cell3Run3 = cell3Paragraph.createRun();
    	                      cell3Run3.setFontSize(18);
    	                      cell3Run3.setText(data[1]);
    	                  }
    	              }

    	              // Save the document
    	              
    	      document.write(out);
    	      out.close();
    	      System.out.println("createdocument.docx written successully");

    	      }   

 
   
    public static void main(String[] args) {
        new ResumeBuilder();
    }
}
