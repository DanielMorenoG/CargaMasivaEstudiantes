-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 05-03-2026 a las 15:26:08
-- Versión del servidor: 10.4.32-MariaDB
-- Versión de PHP: 8.2.12

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `notas2026`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `estudiantes`
--

CREATE TABLE `estudiantes` (
  `id` int(11) NOT NULL,
  `Nombre` varchar(100) DEFAULT NULL,
  `Edad` int(11) DEFAULT NULL,
  `Carrera` varchar(100) DEFAULT NULL,
  `nota1` decimal(3,2) DEFAULT NULL,
  `nota2` decimal(3,2) DEFAULT NULL,
  `nota3` decimal(3,2) DEFAULT NULL,
  `Promedio` decimal(3,2) DEFAULT NULL,
  `Desempeño` varchar(20) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `estudiantes`
--

INSERT INTO `estudiantes` (`id`, `Nombre`, `Edad`, `Carrera`, `nota1`, `nota2`, `nota3`, `Promedio`, `Desempeño`) VALUES
(1, 'Paula', 21, 'Fisica', 4.00, 4.00, 3.00, 3.00, 'Bueno'),
(2, 'Ana', 18, 'Ingenieria', 2.00, 5.00, 3.00, 3.00, 'Regular'),
(3, 'Maria', 23, 'Ingenieria', 5.00, 4.00, 3.00, 4.00, 'Bueno'),
(4, 'Luis', 22, 'Matematicas', 2.00, 3.00, 4.00, 3.00, 'Regular'),
(5, 'Ana', 21, 'Ingenieria', 5.00, 5.00, 5.00, 5.00, 'Excelente'),
(6, 'Maria', 23, 'Ingenieria', 4.00, 3.00, 3.00, 3.00, 'Regular'),
(7, 'Ana', 20, 'Fisica', 2.00, 3.00, 3.00, 2.00, 'Regular'),
(8, 'Luis', 20, 'Ingenieria', 4.00, 2.00, 4.00, 3.00, 'Bueno'),
(9, 'Luis', 23, 'Fisica', 4.00, 3.00, 2.00, 3.00, 'Regular'),
(10, 'Luis', 22, 'Ingenieria', 3.00, 3.00, 2.00, 2.00, 'Regular'),
(11, 'Ana', 20, 'Fisica', 5.00, 3.00, 2.00, 3.00, 'Regular'),
(12, 'Carlos', 19, 'Fisica', 4.00, 2.00, 2.00, 2.00, 'Regular'),
(13, 'Luis', 21, 'Fisica', 2.00, 5.00, 5.00, 4.00, 'Bueno'),
(14, 'Maria', 22, 'Fisica', 5.00, 2.00, 2.00, 3.00, 'Regular'),
(15, 'Jose', 18, 'Fisica', 4.00, 3.00, 2.00, 3.00, 'Regular'),
(16, 'Paula', 21, 'Fisica', 5.00, 4.00, 2.00, 3.00, 'Bueno'),
(17, 'Luis', 22, 'Ingenieria', 2.00, 3.00, 2.00, 2.00, 'Deficiente'),
(18, 'Maria', 22, 'Matematicas', 5.00, 5.00, 2.00, 4.00, 'Bueno'),
(19, 'Luis', 20, 'Matematicas', 5.00, 4.00, 5.00, 4.00, 'Excelente'),
(20, 'Ana', 22, 'Ingenieria', 2.00, 3.00, 4.00, 3.00, 'Regular'),
(21, 'Ana', 20, 'Fisica', 3.00, 5.00, 2.00, 3.00, 'Bueno'),
(22, 'Carlos', 20, 'Ingenieria', 2.00, 4.00, 5.00, 3.00, 'Bueno'),
(23, 'Luis', 23, 'Fisica', 3.00, 2.00, 5.00, 3.00, 'Bueno'),
(24, 'Ana', 21, 'Ingenieria', 3.00, 5.00, 4.00, 4.00, 'Bueno'),
(25, 'Carlos', 19, 'Matematicas', 4.00, 3.00, 3.00, 3.00, 'Regular'),
(26, 'Maria', 18, 'Fisica', 3.00, 3.00, 5.00, 3.00, 'Bueno'),
(27, 'Carlos', 22, 'Matematicas', 3.00, 4.00, 2.00, 3.00, 'Regular'),
(28, 'Luis', 21, 'Ingenieria', 2.00, 3.00, 5.00, 3.00, 'Regular'),
(29, 'Jose', 20, 'Matematicas', 4.00, 3.00, 3.00, 3.00, 'Regular'),
(30, 'Ana', 19, 'Ingenieria', 3.00, 2.00, 3.00, 3.00, 'Regular'),
(31, 'Maria', 18, 'Ingenieria', 5.00, 4.00, 4.00, 4.00, 'Excelente'),
(32, 'Maria', 23, 'Fisica', 2.00, 3.00, 4.00, 3.00, 'Regular'),
(33, 'Jose', 18, 'Matematicas', 5.00, 3.00, 4.00, 4.00, 'Bueno'),
(34, 'Ana', 18, 'Matematicas', 5.00, 2.00, 5.00, 4.00, 'Bueno'),
(35, 'Jose', 20, 'Fisica', 2.00, 2.00, 2.00, 2.00, 'Deficiente'),
(36, 'Paula', 23, 'Fisica', 5.00, 5.00, 3.00, 4.00, 'Bueno'),
(37, 'Jose', 18, 'Fisica', 4.00, 3.00, 3.00, 3.00, 'Bueno'),
(38, 'Ana', 21, 'Fisica', 5.00, 3.00, 5.00, 4.00, 'Excelente'),
(39, 'Ana', 22, 'Ingenieria', 3.00, 5.00, 2.00, 3.00, 'Bueno'),
(40, 'Ana', 23, 'Fisica', 3.00, 2.00, 2.00, 2.00, 'Regular'),
(41, 'Luis', 18, 'Fisica', 3.00, 3.00, 4.00, 3.00, 'Bueno'),
(42, 'Luis', 23, 'Fisica', 3.00, 4.00, 3.00, 3.00, 'Bueno'),
(43, 'Ana', 22, 'Fisica', 2.00, 5.00, 3.00, 3.00, 'Regular'),
(44, 'Maria', 21, 'Fisica', 5.00, 5.00, 4.00, 4.00, 'Excelente'),
(45, 'Paula', 20, 'Matematicas', 2.00, 3.00, 2.00, 2.00, 'Deficiente'),
(46, 'Ana', 18, 'Fisica', 5.00, 2.00, 2.00, 3.00, 'Regular'),
(47, 'Felipe Tellez', 23, 'Ingenieria', 4.00, 4.00, 4.00, 4.00, 'Bueno');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `usuarios`
--

CREATE TABLE `usuarios` (
  `id` int(11) NOT NULL,
  `username` varchar(50) DEFAULT NULL,
  `password` varchar(100) DEFAULT NULL,
  `rol` varchar(20) DEFAULT NULL,
  `carrera` varchar(100) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `usuarios`
--

INSERT INTO `usuarios` (`id`, `username`, `password`, `rol`, `carrera`) VALUES
(1, 'admin', 'aaa', 'administrador', NULL),
(2, 'docente1', 'abcd', 'docente', 'Ingeniería'),
(3, 'docente2', 'abcd', 'docente', 'Administración');

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `estudiantes`
--
ALTER TABLE `estudiantes`
  ADD PRIMARY KEY (`id`);

--
-- Indices de la tabla `usuarios`
--
ALTER TABLE `usuarios`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `username` (`username`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `estudiantes`
--
ALTER TABLE `estudiantes`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=48;

--
-- AUTO_INCREMENT de la tabla `usuarios`
--
ALTER TABLE `usuarios`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
