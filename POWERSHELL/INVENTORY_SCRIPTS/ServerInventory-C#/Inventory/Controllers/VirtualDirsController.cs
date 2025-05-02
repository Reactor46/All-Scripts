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
    public class VirtualDirsController : ApiController
    {
        private InventoryContext db = new InventoryContext();

        // GET: api/VirtualDirs
        public IQueryable<VirtualDir> GetVirtualDirectories()
        {
            return db.VirtualDirectories;
        }

        // GET: api/VirtualDirs/5
        [ResponseType(typeof(VirtualDir))]
        public IHttpActionResult GetVirtualDir(int id)
        {
            VirtualDir virtualDir = db.VirtualDirectories.Find(id);
            if (virtualDir == null)
            {
                return NotFound();
            }

            return Ok(virtualDir);
        }

        // PUT: api/VirtualDirs/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutVirtualDir(int id, VirtualDir virtualDir)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != virtualDir.Id)
            {
                return BadRequest();
            }

            db.Entry(virtualDir).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!VirtualDirExists(id))
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

        // POST: api/VirtualDirs
        [ResponseType(typeof(VirtualDir))]
        public IHttpActionResult PostVirtualDir(VirtualDir virtualDir)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.VirtualDirectories.Add(virtualDir);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = virtualDir.Id }, virtualDir);
        }

        // DELETE: api/VirtualDirs/5
        [ResponseType(typeof(VirtualDir))]
        public IHttpActionResult DeleteVirtualDir(int id)
        {
            VirtualDir virtualDir = db.VirtualDirectories.Find(id);
            if (virtualDir == null)
            {
                return NotFound();
            }

            db.VirtualDirectories.Remove(virtualDir);
            db.SaveChanges();

            return Ok(virtualDir);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool VirtualDirExists(int id)
        {
            return db.VirtualDirectories.Count(e => e.Id == id) > 0;
        }
    }
}