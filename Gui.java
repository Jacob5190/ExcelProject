import javax.swing.*;
import java.awt.*;

public class Gui {
    public static JFrame frame = new JFrame("Email");
    public static JPanel panel = new JPanel();
    public static JButton process = new JButton("开始运行");
    public static JTextArea textOut = new JTextArea();
    public static JScrollPane jScrollPane = new JScrollPane(textOut);
    public static JPanel infoPanel = new JPanel();
    public static JTextField email = new JTextField(15);
    public static JTextField password = new JTextField(15);
    public static JTextField SMTP = new JTextField(15);
    Gui(){
        JLabel emailL = new JLabel("Email: ");
        JLabel passwordL = new JLabel("Password: ");
        JLabel SMTPL = new JLabel("SMTP: ");
        GridBagConstraints gbc = new GridBagConstraints();
        infoPanel.setLayout(new GridBagLayout());
        process.addActionListener(e -> {
            try {
                new SwingWorker(){
                    @Override
                    protected Object doInBackground() throws Exception {
                        process.setEnabled(false);
                        String emailT = email.getText();
                        String passwordT = password.getText();
                        String SMTPT = SMTP.getText();
                        if (emailT.equals("") || passwordT.equals("") || SMTPT.equals("")){
                            JOptionPane.showMessageDialog(null, "请正确填写邮箱信息", "Alert!", JOptionPane.INFORMATION_MESSAGE);
                        }
                        else {
                            textOut.append("Progress started\n");
                            sendMail.myEmailAccount = emailT;
                            textOut.append("Email address set to: " + emailT + "\n");
                            sendMail.myEmailPassword = passwordT;
                            textOut.append("Email password set to: " + passwordT + "\n");
                            sendMail.myEmailSMTPHost = SMTPT;
                            textOut.append("SMTP address set to: " + SMTPT + "\n");
                            Class.function();
                            Thread.sleep(1000);
                            JOptionPane.showMessageDialog(null, "邮件已发送完毕", "Process finished", JOptionPane.INFORMATION_MESSAGE);
                            process.setEnabled(true);
                        }
                        return null;
                    }
                }.execute();
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(null, "出现未知问题,程序即将退出", "Unexpected Error", JOptionPane.ERROR_MESSAGE);
                try {
                    Thread.sleep(5000);
                } catch (InterruptedException exc) {
                    exc.printStackTrace();
                }
                System.exit(0);
            }
        });
        infoPanel.setPreferredSize(new Dimension(270,200));
        gbc.ipady = 5;
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0.3;
        infoPanel.add(emailL, gbc);
        gbc.gridy = 1;
        infoPanel.add(passwordL, gbc);
        gbc.gridy = 2;
        infoPanel.add(SMTPL, gbc);
        gbc.gridx = 1;
        gbc.gridy = 0;
        gbc.weightx = 0.7;
        infoPanel.add(email, gbc);
        gbc.gridy = 1;
        infoPanel.add(password, gbc);
        gbc.gridy = 2;
        infoPanel.add(SMTP, gbc);
        infoPanel.setBorder(BorderFactory.createEmptyBorder(0,10,70,5));
        textOut.setEditable(false);
        jScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        jScrollPane.setBorder(BorderFactory.createEmptyBorder(10,0,5,2));
        frame.setPreferredSize(new Dimension(570, 280));
        frame.setResizable(false);
        frame.setLayout(new BorderLayout());
        panel.add(process);
        frame.add(panel, BorderLayout.SOUTH);
        frame.add(jScrollPane, BorderLayout.CENTER);
        frame.add(infoPanel, BorderLayout.WEST);
        frame.setVisible(true);
        frame.pack();
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
    }

    public static void main(String[] args) {
        new Gui();
    }
}
