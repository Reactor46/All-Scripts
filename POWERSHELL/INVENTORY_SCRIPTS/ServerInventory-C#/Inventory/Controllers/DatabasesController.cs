using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Description;
using Inventory;
using Inventory.Model;

namespace Inventory.Server.Controllers
{
    public class DatabasesController : ApiController
    {
        private InventoryContext db = new InventoryContext();

        // GET: api/Databases
        public IQueryable<Model.Database> GetEnrollments()
        {
            return db.Databases;
        }

        // GET: api/Databases/5
        [ResponseType(typeof(Model.Database))]
        public IHttpActionResult GetDatabase(int id)
        {
            Model.Database database = db.Databases.Find(id);
            if (database == null)
            {
                return NotFound();
            }

            return Ok(database);
        }

        // PUT: api/Databases/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutDatabase(int id, Model.Database database)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != database.Id)
            {
                return BadRequest();
            }

            db.Entry(database).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!DatabaseExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/Databases
        [ResponseType(typeof(Model.Database))]
        public IHttpActionResult PostDatabase(Model.Database database)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.Databases.Add(database);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = database.Id }, database);
        }

        // DELETE: api/Databases/5
        [ResponseType(typeof(Model.Database))]
        public IHttpActionResult DeleteDatabase(int id)
        {
            Model.Database database = db.Databases.Find(id);
            if (database == null)
            {
                return NotFound();
            }

            db.Databases.Remove(database);
            db.SaveChanges();

            return Ok(database);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool DatabaseExists(int id)
        {
            return db.Databases.Count(e => e.Id == id) > 0;
        }
    }
}