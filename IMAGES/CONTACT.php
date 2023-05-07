<link rel="stylesheet" href="style.css">
<section class="about" id="about">
<?php

require_once 'PHPExcel/Classes/PHPExcel.php';


if(isset($_POST['send'])) {
   $name = $_POST['name'];
   $email = $_POST['email'];
   $number = isset($_POST['number']) ? $_POST['number'] : ""; 
   $message = $_POST['message'];

   $objPHPExcel = PHPExcel_IOFactory::load("Contacts.xlsx");
   $worksheet = $objPHPExcel->getActiveSheet();
   $row = $worksheet->getHighestRow() + 1;
   $worksheet->setCellValue('A'.$row, $name);
  $worksheet->setCellValue('B'.$row, $email);
$worksheet->setCellValue('D'.$row, $message);


   $writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
   $writer->save("Contacts.xlsx");
      echo "<script>alert('Votre message a ete envoye avec succes.');</script>";

   
}

?>


 <section class="contact" id="contact">

   <h1 class="heading" data-aos="fade-up"> <span>Contactez-moi</span> </h1>

   <form action="<?php echo htmlspecialchars($_SERVER['PHP_SELF']); ?>" method="post">
      <div class="flex">
         <input data-aos="fade-right" type="text" name="name" placeholder="entrez votre Nom et Prenom" class="box" required>
         <input data-aos="fade-left" type="email" name="email" placeholder="entrez votre e-mail email" class="box" required>
      </div>
      <textarea data-aos="fade-up" name="message" class="box" required placeholder="entrez votre message" cols="30" rows="10"></textarea>
      <input type="submit" data-aos="zoom-in" value="envoyer" name="send" class="btn">
   </form>

   

</section> 