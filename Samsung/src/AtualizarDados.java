/**
 * Create by: Roberto Guimarães Morati Junior
 * Java SE 8, 2015
 * Version 1
 * 
 * Interface
 */
import java.awt.AWTException;
import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Robot;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JButton;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

import java.awt.Window.Type;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;

import javax.swing.text.Document;

import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.InputMethodListener;
import java.awt.event.InputMethodEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;

public class AtualizarDados {

	private JFrame frmCexel;
	private JTextField textFieldSheet;
	private JButton btnNewButton;
	private JLabel lblVisto;
	private JTextField textFieldVisto;
	private JLabel lblSerial;
	private JTextField textFieldSerial;
	private JLabel lblColVisto;
	private JTextField textFieldColVisto;
	private JLabel lblColStatus;
	private JTextField textFieldColStatus;
	private String fileName;
	private String fileShortName;
	private boolean isValid = false;
	private boolean isValidData = false;
	private boolean isValidColVisto = false;
	private boolean isValidColStatus = false;
	private boolean isValidBolColVisto = false;
	private boolean isValidBolColStatus = false;
	private JLabel lblNewLabel_1;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					AtualizarDados window = new AtualizarDados();
					window.frmCexel.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */

	public void setFileShortName(String fileShortName) {
		this.fileShortName = fileShortName;
	}

	public String getFileShortName() {
		return this.fileShortName;
	}

	public AtualizarDados() {
		initialize();
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getFileName() {
		return this.fileName;
	}

	public void setIsValid(boolean v) {
		this.isValid = v;
	}

	public boolean getIsValid() {
		return this.isValid;
	}

	public void setIsValidData(boolean v) {
		this.isValidData = v;
	}

	public boolean getIsValidData() {
		return this.isValidData;
	}

	public void setIsValidColVisto(boolean v) {
		this.isValidColVisto = v;
	}

	public boolean getIsValidColVisto() {
		return this.isValidColVisto;
	}

	public void setIsValidColStatus(boolean v) {
		this.isValidColStatus = v;
	}

	public boolean getIsValidColStatus() {
		return this.isValidColStatus;
	}

	/*
	 * Get the extension of a file.
	 */
	public static String getExtension(File f) {
		String ext = null;
		String s = f.getName();
		int i = s.lastIndexOf('.');

		if (i > 0 && i < s.length() - 1) {
			ext = s.substring(i + 1).toLowerCase();
		}
		return ext;
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmCexel = new JFrame();
		frmCexel.setTitle("CExcel ");
		frmCexel.setType(Type.UTILITY);
		frmCexel.setResizable(false);
		frmCexel.setBounds(80, 80, 376, 292);
		frmCexel.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmCexel.getContentPane().setLayout(null);

		JLabel lblNewLabel = new JLabel("Planilha  ");
		lblNewLabel.setBounds(10, 42, 111, 14);
		frmCexel.getContentPane().add(lblNewLabel);

		textFieldSheet = new JTextField();
		textFieldSheet.setBounds(127, 39, 233, 20);

		frmCexel.getContentPane().add(textFieldSheet);
		textFieldSheet.setColumns(34);

		btnNewButton = new JButton("Atualizar Planilha");
		btnNewButton.setBounds(191, 227, 156, 23);
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				if (getIsValid() == true) {

					String sheet = "";
					try {
						sheet = (String) textFieldSheet.getText();
						textFieldSheet.setForeground(Color.blue);
						setIsValidData(true);
					} catch (Exception e) {
						setIsValidData(false);
						textFieldSheet.setForeground(Color.red);
					}

					String visto = "";
					try {
						visto = (String) textFieldVisto.getText();
						textFieldVisto.setForeground(Color.blue);
						setIsValidData(true);
					} catch (Exception e) {
						setIsValidData(false);
						textFieldVisto.setForeground(Color.red);
					}

					String serial = "";
					try {
						serial = (String) textFieldSerial.getText();
						textFieldSerial.setForeground(Color.blue);
						setIsValidData(true);

					} catch (Exception e) {
						setIsValidData(false);
						textFieldSerial.setForeground(Color.red);
					}

					int colVisto = 0, colStatus = 0;
					try {

						colVisto = Integer.parseInt(textFieldColVisto.getText());
						colVisto = colVisto - 1;

						textFieldColVisto.setForeground(Color.blue);
						setValidBolColVisto(true);
						if (colVisto == 0 || colVisto < 0) {
							setValidBolColVisto(false);
							textFieldColVisto.setForeground(Color.red);
							JOptionPane
									.showMessageDialog(frmCexel,
											"Valor Coluna Visto precisa ser maior que 1!");
						}
					} catch (IllegalArgumentException e) {
						setValidBolColVisto(false);
						textFieldColVisto.setForeground(Color.red);
						JOptionPane.showMessageDialog(frmCexel,
								"Valor Coluna Visto inválido.");
					} catch (Exception e) {
						setValidBolColVisto(false);
						textFieldColVisto.setForeground(Color.red);
						JOptionPane.showMessageDialog(frmCexel,
								"Valor Coluna Visto inválido.");
					}

					try {
						colStatus = Integer.parseInt(textFieldColStatus
								.getText());
						colStatus = colStatus - 1;

						textFieldColStatus.setForeground(Color.blue);
						setValidBolColStatus(true);
						if (colStatus == 0 || colStatus < 0) {
							setValidBolColStatus(false);
							textFieldColStatus.setForeground(Color.red);
							JOptionPane
									.showMessageDialog(frmCexel,
											"Valor Coluna Status precisa ser maior que 1!");
						}
					} catch (IllegalArgumentException e) {
						setValidBolColStatus(false);
						textFieldColStatus.setForeground(Color.red);
						JOptionPane.showMessageDialog(frmCexel,
								"Valor Coluna Status inválido.");
					} catch (Exception e) {
						setValidBolColStatus(false);
						textFieldColStatus.setForeground(Color.red);
						JOptionPane.showMessageDialog(frmCexel,
								"Valor Coluna Status inválido!");
					}

					if (isValidBolColStatus() == true
							&& isValidBolColVisto() == true) {
						if (colStatus > colVisto) {
							setValidBolColStatus(false);
							textFieldColStatus.setForeground(Color.red);
							JOptionPane
									.showMessageDialog(frmCexel,
											"Valor Coluna Status deve ser menor que Valor Coluna Visto.");
						}
					}

					if (getIsValidData() == true) {
						UpdateExcel excel = new UpdateExcel();

						// valid sheet
						if (excel.validateSheet(sheet, getFileName()) == true) {
							textFieldSheet.setForeground(Color.blue);
							setIsValidData(true);

						} else {
							textFieldSheet.setForeground(Color.red);
							setIsValidData(false);
							String msg = "O arquivo " + getFileShortName()
									+ " não possui a planilha: " + sheet + ".";
							JOptionPane.showMessageDialog(frmCexel, msg);
						}
						if (isValidBolColStatus == true
								&& isValidBolColVisto == true) {
							if (getIsValidData() == true) {
								if (excel.validateColVisto(sheet,
										getFileName(), colVisto) == true) {
									textFieldColVisto.setForeground(Color.blue);
									setIsValidColVisto(true);
								} else {
									textFieldColVisto.setForeground(Color.red);
									setIsValidColVisto(false);
									String msg = "Valor Coluna Visto maior que o número de colunas da planilha: "
											+ sheet + ".";
									JOptionPane
											.showMessageDialog(frmCexel, msg);
								}

								if (excel.validateColStatus(sheet,
										getFileName(), colStatus) == true) {
									textFieldColStatus
											.setForeground(Color.blue);
									setIsValidColStatus(true);
								} else {
									textFieldColStatus.setForeground(Color.red);
									setIsValidColStatus(false);
									String msg = "Valor Coluna Status maior que o número de colunas da planilha: "
											+ sheet + ".";
									JOptionPane
											.showMessageDialog(frmCexel, msg);
								}
							}
						}

						if (getIsValidColVisto() == true
								&& getIsValidColStatus() == true
								&& getIsValidData() == true) {

							excel = new UpdateExcel(sheet, getFileName(),
									visto, serial, colVisto, colStatus);

							if (excel.updateSheet() == true) {
								lblNewLabel_1.setText("Encontrado!");
								lblNewLabel_1.setForeground(Color.blue);
							} else {
								lblNewLabel_1.setText("Não encontrado!");
								lblNewLabel_1.setForeground(Color.red);
							}
						}

					}
				} else {
					JOptionPane.showMessageDialog(frmCexel,
							"Selecione um arquivo para atualizar!");
				}
			}
		});

		lblVisto = new JLabel("Visto        ");
		lblVisto.setBounds(10, 67, 117, 14);
		frmCexel.getContentPane().add(lblVisto);

		textFieldVisto = new JTextField();
		textFieldVisto.setBounds(127, 64, 233, 20);
		textFieldVisto.setColumns(34);
		frmCexel.getContentPane().add(textFieldVisto);

		lblSerial = new JLabel("Serial       ");
		lblSerial.setBounds(10, 91, 115, 14);
		frmCexel.getContentPane().add(lblSerial);

		textFieldSerial = new JTextField();
		textFieldSerial.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				int keyCode = e.getKeyCode();
				switch( keyCode ) {
				 case KeyEvent.VK_ENTER:
					 textFieldSerial.selectAll();
					 break;
				}
					
				
			}
		});
	
		DocumentListener documentListener = new DocumentListener() {
			@Override
			public void changedUpdate(DocumentEvent documentEvent) {
				System.out.println("changed");
			}

			@Override
			public void insertUpdate(DocumentEvent documentEvent) {
				// TODO Auto-generated method stub
				System.out.println("Insert" + textFieldSerial.getText());
				try {
					 
					 Robot robot = new Robot();
					
					 robot.keyPress(KeyEvent.VK_ENTER);
					 
					 robot = null;
				} catch (AWTException e) {
					// TODO Auto-generated catch block
					
				}

				
			}

			@Override
			public void removeUpdate(DocumentEvent documentEvent) {
				// TODO Auto-generated method stub
				System.out.println("remove");
			}

			private void printIt(DocumentEvent documentEvent) {
				DocumentEvent.EventType type = documentEvent.getType();
				String typeString = null;
				if (type.equals(DocumentEvent.EventType.CHANGE)) {
					typeString = "Change";
				} else if (type.equals(DocumentEvent.EventType.INSERT)) {
					typeString = "Insert";
					
				} else if (type.equals(DocumentEvent.EventType.REMOVE)) {
					typeString = "Remove";
				}
				System.out.print("Type : " + typeString);
				Document source = documentEvent.getDocument();
				int length = source.getLength();
				System.out.println("Length: " + length);
			}		
		};
		
		textFieldSerial.getDocument().addDocumentListener(documentListener);
		textFieldSerial.setBounds(127, 89, 233, 20);
		textFieldSerial.setColumns(34);
		frmCexel.getContentPane().add(textFieldSerial);

		lblColVisto = new JLabel("Coluna Visto  ");
		lblColVisto.setBounds(8, 140, 119, 14);
		frmCexel.getContentPane().add(lblColVisto);

		textFieldColVisto = new JTextField();
		textFieldColVisto.setBounds(127, 137, 233, 20);
		textFieldColVisto.setColumns(34);
		frmCexel.getContentPane().add(textFieldColVisto);

		lblColStatus = new JLabel("Coluna Status");
		lblColStatus.setBounds(10, 116, 117, 14);
		frmCexel.getContentPane().add(lblColStatus);

		textFieldColStatus = new JTextField();
		textFieldColStatus.setBounds(127, 113, 233, 20);
		textFieldColStatus.setColumns(34);
		frmCexel.getContentPane().add(textFieldColStatus);
		frmCexel.getContentPane().add(btnNewButton);

		JLabel label = new JLabel("");
		label.setBounds(10, 193, 171, 23);
		frmCexel.getContentPane().add(label);

		JButton btnNewButton_1 = new JButton("Selecione o Arquivo");
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ae) {
				JFileChooser fileChooser = new JFileChooser();
				int returnValue = fileChooser.showOpenDialog(null);
				if (returnValue == JFileChooser.APPROVE_OPTION) {
					File selectedFile = fileChooser.getSelectedFile();
					setFileName(selectedFile.getPath());

					String ext = getExtension(selectedFile);
					if (ext.length() > 1) {
						if (ext.equals("xls")) {
							setIsValid(true);
							label.setText(selectedFile.getName());
							setFileShortName(selectedFile.getName());
							label.setForeground(Color.blue);
						} else {
							setIsValid(false);
							label.setText(selectedFile.getName());
							label.setForeground(Color.red);
						}
					} else {
						JOptionPane.showMessageDialog(frmCexel,
								"Selecione um arquivo!");
					}

				}
			}
		});
		btnNewButton_1.setBounds(191, 193, 156, 23);
		frmCexel.getContentPane().add(btnNewButton_1);

		lblNewLabel_1 = new JLabel("");
		lblNewLabel_1.setBounds(10, 227, 184, 18);
		frmCexel.getContentPane().add(lblNewLabel_1);
		lblNewLabel_1.setText("");
		lblNewLabel_1.setForeground(Color.blue);

	}

	public boolean isValidBolColVisto() {
		return isValidBolColVisto;
	}

	public void setValidBolColVisto(boolean isValidBolColVisto) {
		this.isValidBolColVisto = isValidBolColVisto;
	}

	public boolean isValidBolColStatus() {
		return isValidBolColStatus;
	}

	public void setValidBolColStatus(boolean isValidBolColStatus) {
		this.isValidBolColStatus = isValidBolColStatus;
	}
}
