using Microsoft.AspNetCore.Mvc;
using ExcelFilterApi.Models;
using ExcelFilterApi.Services;
using ExcelFilterApi.Data;
using BCrypt.Net;
namespace ExcelFilterApi.Controllers
{
    [ApiController]
    [Route("api/auth")]
    public class AuthController : ControllerBase
    {
        private readonly JwtService _jwtService;

        public AuthController(JwtService jwtService)
        {
            _jwtService = jwtService;
        }

        [HttpPost("register")]
        public IActionResult Register([FromBody] UserDto dto)
        {
            if (UserRepository.Users.Any(u => u.Email == dto.Email))
                return BadRequest("User already exists");

            var user = new User
            {
                Id = UserRepository.Users.Count + 1,
                Email = dto.Email,
                PasswordHash = BCrypt.Net.BCrypt.HashPassword(dto.Password)
            };

            UserRepository.Users.Add(user);
            return Ok("Registered");
        }

        [HttpPost("login")]
        public IActionResult Login([FromBody] UserDto dto)
        {
            var user = UserRepository.Users.FirstOrDefault(u => u.Email == dto.Email);
            if (user == null)
                return Unauthorized();

            if (!BCrypt.Net.BCrypt.Verify(dto.Password, user.PasswordHash))
                return Unauthorized();

            var token = _jwtService.GenerateToken(user);

            Response.Cookies.Append("token", token, new CookieOptions
            {
                HttpOnly = true,
                Secure = true, // בפרודקשן
                SameSite = SameSiteMode.None,
                Expires = DateTime.UtcNow.AddDays(1)
            });

            return Ok("Logged in");
        }

        [HttpPost("logout")]
        public IActionResult Logout()
        {
            Response.Cookies.Delete("token");
            return Ok("Logged out");
        }

        [HttpGet("me")]
        [Microsoft.AspNetCore.Authorization.Authorize]
        public IActionResult Me()
        {
            return Ok(new
            {
                Email = User.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value
            });
        }
    }
}