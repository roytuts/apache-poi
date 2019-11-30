package com.roytuts.generic.way.to.read.excel.apache.poi.model;

import java.time.LocalDate;

public class Profit {

	private LocalDate date;
	private double profit;

	public LocalDate getDate() {
		return date;
	}

	public void setDate(LocalDate date) {
		this.date = date;
	}

	public double getProfit() {
		return profit;
	}

	public void setProfit(double profit) {
		this.profit = profit;
	}

}
