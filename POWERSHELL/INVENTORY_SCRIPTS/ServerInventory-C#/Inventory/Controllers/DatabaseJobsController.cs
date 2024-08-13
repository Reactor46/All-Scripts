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
    public class DatabaseJobsController : ApiController
    {
        private InventoryContext db = new InventoryContext();

        // GET: api/DatabaseJobs
        public IQueryable<DatabaseJob> GetCourses()
        {
            return db.DatabaseJobs;
        }

        // GET: api/DatabaseJobs/5
        [ResponseType(typeof(DatabaseJob))]
        public IHttpActionResult GetDatabaseJob(int id)
        {
            DatabaseJob databaseJob = db.DatabaseJobs.Find(id);
            if (databaseJob == null)
            {
                return NotFound();
            }

            return Ok(databaseJob);
        }

        // PUT: api/DatabaseJobs/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutDatabaseJob(int id, DatabaseJob databaseJob)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != databaseJob.Id)
            {
                return BadRequest();
            }

            db.Entry(databaseJob).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!DatabaseJobExists(id))
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

        // POST: api/DatabaseJobs
        [ResponseType(typeof(DatabaseJob))]
        public IHttpActionResult PostDatabaseJob(DatabaseJob databaseJob)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.DatabaseJobs.Add(databaseJob);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = databaseJob.Id }, databaseJob);
        }

        // DELETE: api/DatabaseJobs/5
        [ResponseType(typeof(DatabaseJob))]
        public IHttpActionResult DeleteDatabaseJob(int id)
        {
            DatabaseJob databaseJob = db.DatabaseJobs.Find(id);
            if (databaseJob == null)
            {
                return NotFound();
            }

            db.DatabaseJobs.Remove(databaseJob);
            db.SaveChanges();

            return Ok(databaseJob);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool DatabaseJobExists(int id)
        {
            return db.DatabaseJobs.Count(e => e.Id == id) > 0;
        }
    }
}