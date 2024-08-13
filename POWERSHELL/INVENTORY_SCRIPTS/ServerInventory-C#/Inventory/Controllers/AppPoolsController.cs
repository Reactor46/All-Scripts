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
    public class AppPoolsController : ApiController
    {
        private InventoryContext db = new InventoryContext();

        // GET: api/AppPools
        public IQueryable<AppPool> GetApplicationPools()
        {
            return db.ApplicationPools;
        }

        // GET: api/AppPools/5
        [ResponseType(typeof(AppPool))]
        public IHttpActionResult GetAppPool(int id)
        {
            AppPool appPool = db.ApplicationPools.Find(id);
            if (appPool == null)
            {
                return NotFound();
            }

            return Ok(appPool);
        }

        // PUT: api/AppPools/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutAppPool(int id, AppPool appPool)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != appPool.Id)
            {
                return BadRequest();
            }

            db.Entry(appPool).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!AppPoolExists(id))
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

        // POST: api/AppPools
        [ResponseType(typeof(AppPool))]
        public IHttpActionResult PostAppPool(AppPool appPool)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.ApplicationPools.Add(appPool);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = appPool.Id }, appPool);
        }

        // DELETE: api/AppPools/5
        [ResponseType(typeof(AppPool))]
        public IHttpActionResult DeleteAppPool(int id)
        {
            AppPool appPool = db.ApplicationPools.Find(id);
            if (appPool == null)
            {
                return NotFound();
            }

            db.ApplicationPools.Remove(appPool);
            db.SaveChanges();

            return Ok(appPool);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool AppPoolExists(int id)
        {
            return db.ApplicationPools.Count(e => e.Id == id) > 0;
        }
    }
}