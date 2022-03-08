package trabalhopoo;
import java.util.Scanner;
public class PartidaTester {
	public static void main(String[] args) {
		for(int i=1;i<4;i++)
		{
			Scanner var = new Scanner(System.in); 
			Time timeLocal = new Time(); 
			Time timeVisitante = new Time();
		
			System.out.println("Iniciando as partidas da rodada do Brasileirão.");
			System.out.println("Vai começar a partida número"+i+"\n");
		
			System.out.println("Entre com o Time Local:");
			String time1 = var.nextLine();   
			timeLocal.setNomeTime(time1);
		
			System.out.print("Entre com o Time Visitante:");
			String time2 = var.nextLine();
			timeVisitante.setNomeTime(time2);
		
			System.out.print("Preparando o jogo número"+i+"\n");
			timeLocal.listar();
			System.out.println("\n");
			timeVisitante.listar();
		
			System.out.println("** Pronto para iniciar partida? <Pressione uma tecla> **\n");
			String beg = var.nextLine();
		
			System.out.println(" Inicio de jogo. Partida iniciada !!\n");
		
			Partida partida = new Partida();
		
			partida.iniciarPartida(timeLocal, timeVisitante);
			partida.mostrarResultado();
			partida.finalizarPartida();
			timeLocal.listar();
			timeVisitante.listar();
		}
		System.out.println("\n");
		System.out.println("***********************************");
		System.out.println("***********************************");
		System.out.println("********* FIM DE RODADA ***********");
		System.out.println("***********************************");
		System.out.println("Integrantes do Grupo: Pedro Henrique da Costa Gaspar.");
		System.out.println("***********************************");
		System.out.println("GitHub>>>https://github.com/racercore96/UNILASALLE_POO_2020.1.git");
		System.out.println("***********************************");
		System.out.println("Vídeo YouTube>>>https://www.youtube.com/watch?v=D5KFVjUCpBg");
		}
} 