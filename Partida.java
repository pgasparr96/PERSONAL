package trabalhopoo;
import java.util.Random;
public class Partida {
	
	Time timeVisitante; //*
	Time timeLocal; //*
	int scoreVisitante;
	int scoreLocal;
	int dataPartida;
	
	public Time getTimeVisitante() {
		return timeVisitante;
	}
	public void setTimeVisitante(Time timeVisitante) {
		this.timeVisitante = timeVisitante;
	}
	public Time getTimeLocal() {
		return timeLocal;
	}
	public void setTimeLocal(Time timeLocal) {
		this.timeLocal = timeLocal;
	}
	public int getScoreVisitante() {
		return scoreVisitante;
	}
	public void setScoreVisitante(int scoreVisitante) {
		this.scoreVisitante = scoreVisitante;
	}
	public int getScoreLocal() {
		return scoreLocal;
	}
	public void setScoreLocal(int scoreLocal) {
		this.scoreLocal = scoreLocal;
	}
	public int getDataPartida() {
		return dataPartida;
	}
	public void setDataPartida(int dataPartida) {
		this.dataPartida = dataPartida;
	}
	public void iniciarPartida(Time timeLocal, Time timeVisitante){
		Random random = new Random();
		this.scoreLocal = random.nextInt(10);
		this.scoreVisitante = random.nextInt(10);
		this.timeLocal = timeLocal; //*
		this.timeVisitante = timeVisitante; //*
	}
	
	public void mostrarResultado() {
		
		System.out.println(timeLocal.getNomeTime() + " " + this.scoreLocal + "X" + this.scoreVisitante + " " + timeVisitante.getNomeTime());
	}
	
	public void finalizarPartida() {
		if(this.scoreVisitante == this.scoreLocal)
		{
			timeLocal.setQtdEmpates(1);
			timeVisitante.setQtdEmpates(1);
			timeLocal.setQtdPontos(1);
			timeVisitante.setQtdPontos(1);
		}
		else if(this.scoreVisitante < this.scoreLocal)
		{
			timeLocal.setQtdVitorias(1);
			timeVisitante.setQtdDerrotas(1);
			timeLocal.setQtdPontos(3);
		}
		else
		{
			timeLocal.setQtdDerrotas(1);
			timeVisitante.setQtdVitorias(1);
			timeVisitante.setQtdPontos(3);
		}
	}
}

