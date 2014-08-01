package com.cnepay.checkexcel.ui;

import java.awt.Color;
import java.awt.event.ComponentEvent;
import java.awt.event.ComponentListener;

import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextPane;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;

import com.cnepay.checkexcel.report.CheckResultMessage;
import com.cnepay.checkexcel.report.StatData;

@SuppressWarnings("serial")
public class MessagePanel extends JPanel implements Runnable, ComponentListener {

	private JTextPane textArea = new JTextPane();
	private JScrollPane scrollPane = new JScrollPane(textArea);
	
	private int lineNumber = 0;
	
	public MessagePanel(String startText) {

//		textArea.setRows(rows);
//		textArea.setColumns(columns);
//		
//		textArea.setWrapStyleWord(true);
//		textArea.setLineWrap(true);
		
//		textArea.setText(startText);
		
//		textArea.setPreferredSize(new Dimension(650, 350));
//		textArea.setMaximumSize(new Dimension(800, 400));

//		textArea.setBackground(Color.YELLOW);

		// 内容只读
		textArea.setEditable(false);
		// 初始文字
		textArea.setText(startText);

//		textArea.setSelectionColor(textArea.getBackground());
		
		//scrollPane.setBackground(Color.BLUE);
		
		// 根据需要出现滚动条
		scrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);

		
//		scrollPane.setViewportBorder(BorderFactory.createBevelBorder(BevelBorder.LOWERED));
//		scrollPane.setBorder(BorderFactory.createEtchedBorder());
		
//		scrollPane.setPreferredSize(textArea.getPreferredSize());
//		scrollPane.setMaximumSize(textArea.getMaximumSize());
		
//		scrollPane.validate();
		
//		scrollPane.setBorder(BorderFactory.createEmptyBorder());
		
		// 清除panel的布局
		this.setLayout(null);
		// 添加JScrollPane
		this.add(scrollPane);
		
//		setBorder(BorderFactory.createEtchedBorder());
		
		// 添加Component事件监听，主要是resize
		this.addComponentListener(this);
	}	

	/**
	 * 显示系统消息
	 * @param message
	 */
	public void addSystemMessage(String message) {
		Document docs = textArea.getDocument();
		
		SimpleAttributeSet attrSet = new SimpleAttributeSet();
		StyleConstants.setForeground(attrSet, Color.BLUE);
//		StyleConstants.setUnderline(attrSet, true);
//		StyleConstants.setItalic(attrSet, true);
//		StyleConstants.setFontSize(attrSet, 24);
		
		try {
			docs.insertString(docs.getLength(), "\r\n【系统】" + message, attrSet);
		} catch (BadLocationException e) {
			e.printStackTrace();
		}
		
		textArea.setCaretPosition(docs.getLength());
	}
	
	/**
	 * 显示错误消息
	 * @param message
	 */
	public void addErrorMessage(String message) {
		Document docs = textArea.getDocument();
		
		SimpleAttributeSet attrSet = new SimpleAttributeSet();
		StyleConstants.setForeground(attrSet, Color.RED);

		try {
			docs.insertString(docs.getLength(), "\r\n" + lineNumber++ + ":【错误" + StatData.errorNumber + "】" + message, attrSet);
		} catch (BadLocationException e) {
			e.printStackTrace();
		}
		
		textArea.setCaretPosition(docs.getLength());
	}

	/**
	 * 显示警告消息
	 * @param message
	 */
	public void addWarnMessage(String message) {
		Document docs = textArea.getDocument();
		
		SimpleAttributeSet attrSet = new SimpleAttributeSet();
		StyleConstants.setForeground(attrSet, Color.ORANGE);

		try {
			docs.insertString(docs.getLength(), "\r\n" + lineNumber++ + ":【警告】" + message, attrSet);
		} catch (BadLocationException e) {
			e.printStackTrace();
		}	
		
		textArea.setCaretPosition(docs.getLength());
	}

	/**
	 * 显示正常消息
	 * @param message
	 */
	public void addOkMessage(String message) {
		Document docs = textArea.getDocument();
		
		SimpleAttributeSet attrSet = new SimpleAttributeSet();
		StyleConstants.setForeground(attrSet, Color.BLACK);

		try {
			docs.insertString(docs.getLength(), "\r\n" + lineNumber++ + ": " + message, attrSet);
		} catch (BadLocationException e) {
			e.printStackTrace();
		}
		
		textArea.setCaretPosition(docs.getLength());
	}
	
	/**
	 * 显示校验结果消息，格式CheckResultMessage
	 * @param message
	 */
	public void addMessage(CheckResultMessage message) {
		if (message == null || message.getMessage().trim().isEmpty()) {
			return;
		}
		
		if (message.getType() == CheckResultMessage.SYSTEM) {
			addSystemMessage(message.getMessage());
		} else if (message.getType() == CheckResultMessage.CHECK_OK) {
			this.addOkMessage(message.getMessage());
		} else if (message.getType() == CheckResultMessage.CHECK_ERROR) {
			StatData.errorNumber++;
			this.addErrorMessage(message.getMessage());
		} else if (message.getType() == CheckResultMessage.CHECK_WARN) {
			this.addWarnMessage(message.getMessage());
		}		
	}
	
	/**
	 * Thread: 获取校验消息队列里push的消息，显示消息并从队列移除该消息
	 */
	@Override
	public void run() {
		addSystemMessage("请选择备付金报表存放的文件夹");

		while (true) {
			// 如果队列里有消息，显示并移除
			if (Main1.list.size() > 0) {
				CheckResultMessage message = Main1.list.remove(0);
				addMessage(message);
			}
			
			// 需要sleep
			try {
				Thread.sleep(50);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}			
		}
	}

	/**
	 * 改变窗口大小事件
	 */
	@Override
	public void componentResized(ComponentEvent e) {
		// 根据panel自适应大小，调整scrollpane和textArea的位置
		scrollPane.setBounds(0, 0, this.getWidth(), this.getHeight());
		textArea.setSize(scrollPane.getSize());
		
		// 调整滚动条到随后
		textArea.selectAll();
		// 去掉全选效果
		textArea.setCaretPosition(textArea.getDocument().getLength());
		
//		JScrollBar bar = scrollPane.getVerticalScrollBar();
//		bar.setValue(bar.getMaximum());
//		scrollPane.getVerticalScrollBar().setValue(scrollPane.getVerticalScrollBar().getMaximum());
		
//		System.out.println("componentResized: bounds:" + this.getBounds() + ", size=" + this.getSize());
//		System.out.println("textArea bounds=" + textArea.getBounds() + ", caret pos=" + textArea.getCaretPosition());
//		System.out.println("scrollPane bounds=" + scrollPane.getBounds());
////		System.out.println("bar pos=" + bar.getValue() + ", maxpos=" + bar.getMaximum());
//		System.out.println("========");
	}


	@Override
	public void componentHidden(ComponentEvent e) {
	}

	@Override
	public void componentMoved(ComponentEvent e) {
	}

	@Override
	public void componentShown(ComponentEvent e) {
	}
	
}
