package com.example.demo;

import com.example.demo.objects.Aluno;
import com.example.demo.services.OptionsSpreadSheetService;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static com.example.demo.enums.HeaderName.*;

@SpringBootApplication
public class DemoApplication {
    public static void main(String[] args) {
        SpringApplication.run(DemoApplication.class, args);

/*        String dirName = "C:\\Users\\rafael.serdeiro\\";
        String fileName = "planilha";

        String titleTag = "ALUNOS";

        String xlsContent = "application/vnd.ms-excel";

        String xlsxContent = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        String odsContent = "application/vnd.oasis.opendocument.spreadsheet";

        List<String> headers = new ArrayList<>();

        headers.add(String.valueOf(NOME));
        headers.add(String.valueOf(RA));
        headers.add(String.valueOf(NOTA_1));
        headers.add(String.valueOf(NOTA_2));
        headers.add(String.valueOf(MEDIA));

        List<Aluno> listaDeAlunos = new ArrayList<>();
        listaDeAlunos.add(new Aluno("Eduardo", "9876525", 7, 8, 0, false));
        listaDeAlunos.add(new Aluno("Luiz", "1234466", 5, 8, 0, false));
        listaDeAlunos.add(new Aluno("Bruna", "6545657", 7, 6, 0, false));
        listaDeAlunos.add(new Aluno("Carlos", "3456558", 10, 3, 0, false));
        listaDeAlunos.add(new Aluno("Sonia", "6544546", 7, 8, 0, false));
        listaDeAlunos.add(new Aluno("Brianda", "3234535", 6, 5, 0, false));
        listaDeAlunos.add(new Aluno("Pedro", "4234524", 7, 5, 0, false));
        listaDeAlunos.add(new Aluno("Julio", "5434513", 7, 2, 0, false));
        listaDeAlunos.add(new Aluno("Henrique", "6543452", 7, 8, 0, false));
        listaDeAlunos.add(new Aluno("Fernando", "4345651", 5, 8, 0, false));
        listaDeAlunos.add(new Aluno("Vitor", "4332341", 7, 9, 0, false));

        String[] aluno1 = {
                String.valueOf(listaDeAlunos.get(0).getNome()),
                String.valueOf(listaDeAlunos.get(0).getRa()),
                String.valueOf(listaDeAlunos.get(0).getNota1()),
                String.valueOf(listaDeAlunos.get(0).getNota2()),
                String.valueOf(listaDeAlunos.get(0).getMedia())
        };

        String[] aluno2 = {
                String.valueOf(listaDeAlunos.get(1).getNome()),
                String.valueOf(listaDeAlunos.get(1).getRa()),
                String.valueOf(listaDeAlunos.get(1).getNota1()),
                String.valueOf(listaDeAlunos.get(1).getNota2()),
                String.valueOf(listaDeAlunos.get(1).getMedia())
        };

        String[] aluno3 = {
                String.valueOf(listaDeAlunos.get(2).getNome()),
                String.valueOf(listaDeAlunos.get(2).getRa()),
                String.valueOf(listaDeAlunos.get(2).getNota1()),
                String.valueOf(listaDeAlunos.get(2).getNota2()),
                String.valueOf(listaDeAlunos.get(2).getMedia())};

        String[] aluno4 = {
                String.valueOf(listaDeAlunos.get(3).getNome()),
                String.valueOf(listaDeAlunos.get(3).getRa()),
                String.valueOf(listaDeAlunos.get(3).getNota1()),
                String.valueOf(listaDeAlunos.get(3).getNota2()),
                String.valueOf(listaDeAlunos.get(3).getMedia())};

        String[] aluno5 = {
                String.valueOf(listaDeAlunos.get(4).getNome()),
                String.valueOf(listaDeAlunos.get(4).getRa()),
                String.valueOf(listaDeAlunos.get(4).getNota1()),
                String.valueOf(listaDeAlunos.get(4).getNota2()),
                String.valueOf(listaDeAlunos.get(4).getMedia())};

        String[] aluno6 = {
                String.valueOf(listaDeAlunos.get(5).getNome()),
                String.valueOf(listaDeAlunos.get(5).getRa()),
                String.valueOf(listaDeAlunos.get(5).getNota1()),
                String.valueOf(listaDeAlunos.get(5).getNota2()),
                String.valueOf(listaDeAlunos.get(5).getMedia())};

        String[] aluno7 = {
                String.valueOf(listaDeAlunos.get(6).getNome()),
                String.valueOf(listaDeAlunos.get(6).getRa()),
                String.valueOf(listaDeAlunos.get(6).getNota1()),
                String.valueOf(listaDeAlunos.get(6).getNota2()),
                String.valueOf(listaDeAlunos.get(6).getMedia())};

        String[] aluno8 = {
                String.valueOf(listaDeAlunos.get(7).getNome()),
                String.valueOf(listaDeAlunos.get(7).getRa()),
                String.valueOf(listaDeAlunos.get(7).getNota1()),
                String.valueOf(listaDeAlunos.get(7).getNota2()),
                String.valueOf(listaDeAlunos.get(7).getMedia())};


        List<List<String>> dadosDosAlunos = new ArrayList<>();

        dadosDosAlunos.add(Arrays.asList(aluno1));
        dadosDosAlunos.add(Arrays.asList(aluno2));
        dadosDosAlunos.add(Arrays.asList(aluno3));
        dadosDosAlunos.add(Arrays.asList(aluno4));
        dadosDosAlunos.add(Arrays.asList(aluno5));
        dadosDosAlunos.add(Arrays.asList(aluno6));
        dadosDosAlunos.add(Arrays.asList(aluno7));
        dadosDosAlunos.add(Arrays.asList(aluno8));

        java-api-export-local-universal-spreadsheet

        System.out.println(dadosDosAlunos);

        try {
            OptionsSpreadSheetService.exportSpreadSheet(headers, dadosDosAlunos, dirName, fileName, titleTag, xlsxContent);
        } catch (IOException e) {
            e.printStackTrace();
        }*/
    }

}

