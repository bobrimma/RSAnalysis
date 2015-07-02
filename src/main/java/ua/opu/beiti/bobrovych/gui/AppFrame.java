package ua.opu.beiti.bobrovych.gui;

import java.awt.Component;
import java.awt.Container;
import java.awt.EventQueue;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JTextField;
import javax.swing.UIManager;

import ua.opu.beiti.bobrovych.util.ExcelFilter;
import ua.opu.beiti.bobrovych.analysis.HerstCounter;
import ua.opu.beiti.bobrovych.analysis.VNCounter;
import ua.opu.beiti.bobrovych.exceptions.BlankCellException;
import ua.opu.beiti.bobrovych.exceptions.NotEnoughtDataException;
import ua.opu.beiti.bobrovych.exceptions.TooMuchDataException;
import ua.opu.beiti.bobrovych.util.ExcelParser;

public class AppFrame {
	private JFrame frame = new JFrame();
	private JPanel panel = new JPanel();
	private JRadioButton radioBtnHerst;
	private JRadioButton radioBthVn;
	private JTextField textFieldSourceFile;
	private JButton btnChooseSource;
	private JTextField textFieldSaveToFile;
	private JButton btnChooseDestination;
	private JFileChooser fileChooser;

	public AppFrame() {
		initializeUI();
	}

	private void initializeUI() {
		JLabel label = new JLabel("����� ������ ���������?");
		radioBtnHerst = new JRadioButton("����������� ������");
		radioBtnHerst.setSelected(true);
		radioBthVn = new JRadioButton("Vn");
		ButtonGroup group = new ButtonGroup();
		group.add(radioBtnHerst);
		group.add(radioBthVn);

		panel.setLayout(new GridBagLayout());
		panel.add(label, new GridBagConstraints(0, 0, 1, 1, 0, 0,
				GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5,
						5, 5, 5), 0, 0));
		panel.add(radioBtnHerst, new GridBagConstraints(
				GridBagConstraints.RELATIVE, 0, 1, 1, 0, 0,
				GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5,
						5, 5, 5), 0, 0));
		panel.add(radioBthVn, new GridBagConstraints(
				GridBagConstraints.RELATIVE, 0, 1, 1, 0, 0,
				GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5,
						5, 5, 5), 0, 0));

		label = new JLabel("���� �������� ������:");
		textFieldSourceFile = new JTextField();
		textFieldSourceFile.setEditable(false);
		btnChooseSource = new JButton("�����");
		btnChooseSource.addActionListener(new FileChooserButtonListener());
		panel.add(label, new GridBagConstraints(0, GridBagConstraints.RELATIVE,
				GridBagConstraints.RELATIVE, 1, 0, 0, GridBagConstraints.WEST,
				GridBagConstraints.NONE, new Insets(5, 5, 5, 5), 0, 0));
		panel.add(textFieldSourceFile, new GridBagConstraints(0,
				GridBagConstraints.RELATIVE, GridBagConstraints.RELATIVE, 1, 0,
				0, GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL,
				new Insets(5, 5, 5, 5), 0, 0));
		panel.add(btnChooseSource, new GridBagConstraints(5, 2, 1, 1, 0, 0,
				GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5,
						5, 5, 5), 0, 0));

		label = new JLabel("��������� ������� � ����:");
		textFieldSaveToFile = new JTextField();
		textFieldSaveToFile.setEditable(false);
		btnChooseDestination = new JButton("�����");
		btnChooseDestination.addActionListener(new FileChooserButtonListener());

		panel.add(label, new GridBagConstraints(0, GridBagConstraints.RELATIVE,
				GridBagConstraints.RELATIVE, 1, 0, 0, GridBagConstraints.WEST,
				GridBagConstraints.NONE, new Insets(5, 5, 5, 5), 0, 0));
		panel.add(textFieldSaveToFile, new GridBagConstraints(0,
				GridBagConstraints.RELATIVE, GridBagConstraints.RELATIVE, 1, 0,
				0, GridBagConstraints.WEST, GridBagConstraints.HORIZONTAL,
				new Insets(5, 5, 5, 5), 0, 0));
		panel.add(btnChooseDestination, new GridBagConstraints(5, 4, 1, 1, 0,
				0, GridBagConstraints.WEST, GridBagConstraints.NONE,
				new Insets(5, 5, 5, 5), 0, 0));

		label = new JLabel(
				"��������! ���� ���� ��� ����������,�� ��� ���������� ����� ������������.");

		panel.add(label, new GridBagConstraints(0, GridBagConstraints.RELATIVE,
				GridBagConstraints.REMAINDER, 1, 0, 0, GridBagConstraints.WEST,
				GridBagConstraints.NONE, new Insets(5, 5, 5, 5), 0, 0));

		JButton btnFinish = new JButton("������");
		btnFinish.addActionListener(new FinishButtonListener());
		JButton btnClean = new JButton("��������");
		btnClean.addActionListener(new ClearButtonListener());
		panel.add(btnClean, new GridBagConstraints(4, 6, 1, 1, 0, 0,
				GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5,
						5, 5, 5), 0, 0));
		panel.add(btnFinish, new GridBagConstraints(5, 6, 1, 1, 0, 0,
				GridBagConstraints.WEST, GridBagConstraints.NONE, new Insets(5,
						5, 5, 5), 0, 0));

		frame.setTitle("R/S ������ v.1.0.0");
		frame.add(panel);
		frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		frame.pack();
		frame.setResizable(false);
		frame.setLocationRelativeTo(null);
		frame.setVisible(true);
	}

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					try {
						for (UIManager.LookAndFeelInfo lafInfo : UIManager
								.getInstalledLookAndFeels()) {
							if ("Nimbus".equals(lafInfo.getName())) {
								UIManager.setLookAndFeel(lafInfo.getClassName());
								break;
							}
						}
					} catch (Exception e) {
						try {
							UIManager.setLookAndFeel(UIManager
									.getCrossPlatformLookAndFeelClassName());
						} catch (Exception e1) {
							e1.printStackTrace();
						}
					}
					new AppFrame();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	private void resetFields() {
		textFieldSourceFile.setText("");
		textFieldSaveToFile.setText("");
		radioBtnHerst.setSelected(true);

	}

	private class ClearButtonListener implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent arg0) {
			resetFields();
		}
	}

	private class FinishButtonListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent arg0) {
			if ((!textFieldSourceFile.getText().trim().equals(""))
					&& (!textFieldSaveToFile.getText().trim().equals(""))) {
				ExcelParser jp = new ExcelParser();
				if (radioBtnHerst.isSelected()) {
					try {
						jp.parseExcel(textFieldSourceFile.getText());
						HerstCounter herst = new HerstCounter();

						herst.createStatisticsSheet("����������� ���������� �� ���������� - ����������� ������");
						for (int i = 0; i < jp.getYears().size(); i++) {
							herst.countHerst(jp.getInputData()[i], jp
									.getYears().get(i));
						}
						herst.saveWorkbook(textFieldSaveToFile.getText());

						JOptionPane.showMessageDialog(frame,
								"������ ������� ��������." + "\n"
										+ "���������� ��������� � �����: "
										+ "\n" + textFieldSaveToFile.getText());
						resetFields();
					} catch (NotEnoughtDataException e) {
						JOptionPane.showMessageDialog(frame, e.getMessage(),
								"������ � �������� ������.",
								JOptionPane.ERROR_MESSAGE);
					} catch (TooMuchDataException e) {
						JOptionPane.showMessageDialog(frame, e.getMessage(),
								"������ � �������� ������.",
								JOptionPane.ERROR_MESSAGE);
					} catch (BlankCellException e) {
						JOptionPane.showMessageDialog(frame, e.getMessage(),
								"������ � �������� ������.",
								JOptionPane.ERROR_MESSAGE);
					}
				} else {
					try {
						jp.parseExcelForVn(textFieldSourceFile.getText());
						VNCounter vn = new VNCounter();
						vn.createStatisticsSheet("����������� ���������� �� ���������� - Vn");
						for (int i = 0; i < jp.getYears().size(); i++) {
							vn.countVn(jp.getData().get(i), jp.getYears()
									.get(i), jp.getR().get(i));
						}
						vn.saveWorkbook(textFieldSaveToFile.getText());

						JOptionPane.showMessageDialog(frame,
								"������ ������� ��������." + "\n"
										+ "���������� ��������� � �����: "
										+ "\n" + textFieldSaveToFile.getText());
						resetFields();
					} catch (NotEnoughtDataException e) {
						JOptionPane.showMessageDialog(frame, e.getMessage(),
								"������ � �������� ������.",
								JOptionPane.ERROR_MESSAGE);
					}

				}

			} else {
				JOptionPane.showMessageDialog(frame,
						"�������� ����� �������� � �������� ������.",
						"��������", JOptionPane.WARNING_MESSAGE);
			}
		}
	}

	public class FileChooserButtonListener implements ActionListener {
		public boolean disableTF(Container c) {
			Component[] cmps = c.getComponents();
			for (Component cmp : cmps) {
				if (cmp instanceof JTextField) {
					((JTextField) cmp).setEnabled(false);
					return true;
				}
				if (cmp instanceof Container) {
					if (disableTF((Container) cmp))
						return true;
				}
			}
			return false;
		}

		@Override
		public void actionPerformed(ActionEvent arg0) {
			if (fileChooser == null) {
				fileChooser = new JFileChooser();
				fileChooser.addChoosableFileFilter(new ExcelFilter());
				fileChooser.setAcceptAllFileFilterUsed(false);
				fileChooser.setMultiSelectionEnabled(false);
				disableTF(fileChooser);
			}

			int returnVal = fileChooser
					.showDialog(fileChooser, "�������� ����");

			if (returnVal == JFileChooser.APPROVE_OPTION) {

				File file = fileChooser.getSelectedFile();
				if (arg0.getSource().equals(btnChooseSource)) {
					textFieldSourceFile.setText(file.getAbsolutePath());
					textFieldSaveToFile.setText(file.getParent() + "\\output_"
							+ file.getName());

				} else {
					textFieldSaveToFile.setText(file.getAbsolutePath());
				}

			}

			// Reset the file chooser for the next time it's shown.
			// fc.setSelectedFile(null);

		}

	}
}
