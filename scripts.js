<script>
  const input = document.getElementById('yearInput');

  input.addEventListener('focus', function() {
    if (this.value === 'ทั้งหมด') {
      this.value = ''; // ล้างค่าเมื่อคลิกหรือเริ่มพิมพ์
    }
  });

  input.addEventListener('blur', function() {
    if (this.value.trim() === '') {
      this.value = 'ทั้งหมด'; // ถ้าไม่ได้พิมพ์อะไร กลับไปเป็นค่าเริ่มต้น
    }
  });
</script>