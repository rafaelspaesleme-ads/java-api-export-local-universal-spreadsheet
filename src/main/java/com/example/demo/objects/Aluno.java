package com.example.demo.objects;

public class Aluno {
    private String nome;
    private String ra;
    private double nota1;
    private double nota2;
    private double media;
    private boolean aprovado;

    public Aluno() { }

    public Aluno(String nome, String ra, double nota1, double nota2,
                 double media, boolean aprovado) {
        super();
        this.nome = nome;
        this.ra = ra;
        this.nota1 = nota1;
        this.nota2 = nota2;
        this.media = media;
        this.aprovado = aprovado;
    }



    public String getNome() {
        return nome;
    }
    public void setNome(String nome) {
        this.nome = nome;
    }
    public String getRa() {
        return ra;
    }
    public void setRa(String ra) {
        this.ra = ra;
    }
    public double getNota1() {
        return nota1;
    }
    public void setNota1(double nota1) {
        this.nota1 = nota1;
    }
    public double getNota2() {
        return nota2;
    }
    public void setNota2(double nota2) {
        this.nota2 = nota2;
    }
    public double getMedia() {
        return media;
    }
    public void setMedia(double media) {
        this.media = media;
    }

    public boolean isAprovado() {
        return aprovado;
    }

    public void setAprovado(boolean aprovado) {
        this.aprovado = aprovado;
    }


    public static final class AlunoBuilder {
        private Aluno aluno;

        private AlunoBuilder() {
            aluno = new Aluno();
        }

        public static AlunoBuilder anAluno() {
            return new AlunoBuilder();
        }

        public AlunoBuilder comNome(String nome) {
            aluno.setNome(nome);
            return this;
        }

        public AlunoBuilder comRa(String ra) {
            aluno.setRa(ra);
            return this;
        }

        public AlunoBuilder comNota1(double nota1) {
            aluno.setNota1(nota1);
            return this;
        }

        public AlunoBuilder comNota2(double nota2) {
            aluno.setNota2(nota2);
            return this;
        }

        public AlunoBuilder comMedia(double media) {
            aluno.setMedia(media);
            return this;
        }

        public AlunoBuilder comAprovado(boolean aprovado) {
            aluno.setAprovado(aprovado);
            return this;
        }

        public Aluno build() {
            return aluno;
        }
    }
}
