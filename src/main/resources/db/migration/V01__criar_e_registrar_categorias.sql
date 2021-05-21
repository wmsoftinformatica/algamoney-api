CREATE TABLE categoria (
	codigo BIGINT PRIMARY KEY AUTO_INCREMENT,
	nome VARCHAR(50) NOT NULL
)ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

INSERT INTO categoria (nome) values ('Lazer');
INSERT INTO categoria (nome) values ('Alimentação');
INSERT INTO categoria (nome) values ('Supermercado');
INSERT INTO categoria (nome) values ('Farmacia');
INSERT INTO categoria (nome) values ('Outros'); 