package com.example.demo.objects;

public class Header {
    private String name;

    public Header(String name) {
        this.name = name;
    }

    public Header() {
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public static final class HeaderBuilder {
        private Header header;

        private HeaderBuilder() {
            header = new Header();
        }

        public static HeaderBuilder aHeader() {
            return new HeaderBuilder();
        }

        public HeaderBuilder comName(String name) {
            header.setName(name);
            return this;
        }

        public Header build() {
            return header;
        }
    }
}
